from __future__ import annotations

from collections import defaultdict, deque
from dataclasses import dataclass

from core.models import (
    KiotData,
    PaymentMatchResult,
    PaymentRecord,
    PaymentReconciliation,
    SmartData,
)


@dataclass
class _Edge:
    to: int
    reverse: int
    capacity: int
    cost: int
    pair: tuple[int, int] | None = None


def _add_edge(graph: list[list[_Edge]], source: int, target: int, capacity: int, cost: int, pair=None):
    forward = _Edge(target, len(graph[target]), capacity, cost, pair)
    backward = _Edge(source, len(graph[source]), 0, -cost, None)
    graph[source].append(forward)
    graph[target].append(backward)


def _min_cost_matching(
    smart: list[PaymentRecord],
    kiot: list[PaymentRecord],
    max_delay_days: int,
    forbidden: set[tuple[int, int]] | None = None,
) -> tuple[int, int, list[tuple[int, int]]]:
    forbidden = forbidden or set()
    smart_count = len(smart)
    kiot_count = len(kiot)
    source = 0
    smart_offset = 1
    kiot_offset = smart_offset + smart_count
    sink = kiot_offset + kiot_count
    graph: list[list[_Edge]] = [[] for _ in range(sink + 1)]

    for i in range(smart_count):
        _add_edge(graph, source, smart_offset + i, 1, 0)
    for j in range(kiot_count):
        _add_edge(graph, kiot_offset + j, sink, 1, 0)
    for i, smart_record in enumerate(smart):
        for j, kiot_record in enumerate(kiot):
            if (i, j) in forbidden:
                continue
            delay = (kiot_record.payment_date - smart_record.payment_date).days
            if 0 <= delay <= max_delay_days:
                _add_edge(graph, smart_offset + i, kiot_offset + j, 1, delay, (i, j))

    total_flow = 0
    total_cost = 0
    node_count = len(graph)
    while True:
        distance = [10**12] * node_count
        parent: list[tuple[int, int] | None] = [None] * node_count
        in_queue = [False] * node_count
        distance[source] = 0
        queue = deque([source])
        in_queue[source] = True
        while queue:
            node = queue.popleft()
            in_queue[node] = False
            for edge_index, edge in enumerate(graph[node]):
                if edge.capacity <= 0:
                    continue
                candidate = distance[node] + edge.cost
                if candidate < distance[edge.to]:
                    distance[edge.to] = candidate
                    parent[edge.to] = (node, edge_index)
                    if not in_queue[edge.to]:
                        queue.append(edge.to)
                        in_queue[edge.to] = True
        if parent[sink] is None:
            break
        node = sink
        while node != source:
            previous, edge_index = parent[node]  # type: ignore[misc]
            edge = graph[previous][edge_index]
            edge.capacity -= 1
            graph[node][edge.reverse].capacity += 1
            node = previous
        total_flow += 1
        total_cost += distance[sink]

    pairs: list[tuple[int, int]] = []
    for i in range(smart_count):
        for edge in graph[smart_offset + i]:
            if edge.pair is not None and edge.capacity == 0:
                pairs.append(edge.pair)
    pairs.sort()
    return total_flow, total_cost, pairs


def _matching_is_ambiguous(
    smart: list[PaymentRecord],
    kiot: list[PaymentRecord],
    max_delay_days: int,
    flow: int,
    cost: int,
    pairs: list[tuple[int, int]],
) -> bool:
    for pair in pairs:
        alternative_flow, alternative_cost, _ = _min_cost_matching(
            smart, kiot, max_delay_days, forbidden={pair}
        )
        if alternative_flow == flow and alternative_cost == cost:
            return True
    return False


def _payment_sort_key(record: PaymentRecord):
    time_key = (
        record.payment_time.hour,
        record.payment_time.minute,
        record.payment_time.second,
    ) if record.payment_time else (-1, -1, -1)
    return record.payment_date, time_key, record.source_row


def _classify_unmatched_kiot(
    kiot_record: PaymentRecord,
    same_amount_smart: list[PaymentRecord],
    has_valid_candidate: bool,
) -> tuple[str, str]:
    if has_valid_candidate:
        return "SMART THIẾU", "Không còn giao dịch Smart chưa sử dụng có cùng số tiền và ngày hợp lệ"
    if not same_amount_smart:
        return "SMART THIẾU", "Không tìm thấy giao dịch Smart có cùng số tiền"
    return (
        "SMART THIẾU",
        "Không tìm thấy giao dịch Smart có cùng số tiền trong khoảng ngày hợp lệ",
    )


def reconcile_payments(
    smart_data: SmartData,
    kiot_data: KiotData,
    max_delay_days: int = 2,
) -> PaymentReconciliation:
    issues = list(smart_data.issues) + list(kiot_data.issues)
    smart_by_amount: dict[int, list[PaymentRecord]] = defaultdict(list)
    kiot_by_amount: dict[int, list[PaymentRecord]] = defaultdict(list)
    for record in smart_data.payments:
        smart_by_amount[record.amount].append(record)
    for record in kiot_data.payments:
        kiot_by_amount[record.amount].append(record)

    results: list[PaymentMatchResult] = []
    for amount in sorted(set(smart_by_amount) | set(kiot_by_amount)):
        smart = sorted(smart_by_amount.get(amount, []), key=_payment_sort_key)
        kiot = sorted(kiot_by_amount.get(amount, []), key=_payment_sort_key)
        flow, cost, pairs = _min_cost_matching(smart, kiot, max_delay_days)

        if pairs and _matching_is_ambiguous(smart, kiot, max_delay_days, flow, cost, pairs):
            reason = "Có nhiều cách ghép cùng số tiền và cùng độ lệch ngày; cần kiểm tra thủ công"
            for record in smart:
                results.append(PaymentMatchResult("CẦN KIỂM TRA", reason, smart=record))
            for record in kiot:
                results.append(PaymentMatchResult("CẦN KIỂM TRA", reason, kiot=record))
            continue

        matched_smart = {i for i, _ in pairs}
        matched_kiot = {j for _, j in pairs}
        for i, j in pairs:
            delay = (kiot[j].payment_date - smart[i].payment_date).days
            results.append(
                PaymentMatchResult(
                    status="KHỚP",
                    reason=f"Cùng số tiền; KIOT trễ Smart {delay} ngày",
                    smart=smart[i],
                    kiot=kiot[j],
                    day_difference=delay,
                )
            )

        for j, record in enumerate(kiot):
            if j in matched_kiot:
                continue
            valid_candidate_exists = any(
                0 <= (record.payment_date - candidate.payment_date).days <= max_delay_days
                for candidate in smart
            )
            status, reason = _classify_unmatched_kiot(record, smart, valid_candidate_exists)
            results.append(PaymentMatchResult(status, reason, kiot=record))

        for i, record in enumerate(smart):
            if i not in matched_smart:
                results.append(
                    PaymentMatchResult(
                        status="KIOT THIẾU",
                        reason="Không tìm thấy giao dịch KIOT chưa sử dụng có cùng số tiền và ngày hợp lệ",
                        smart=record,
                    )
                )

    results.sort(
        key=lambda item: (
            (item.smart or item.kiot).payment_date if (item.smart or item.kiot) else 0,
            item.status,
            (item.smart or item.kiot).source_row if (item.smart or item.kiot) else 0,
        )
    )
    return PaymentReconciliation(results=results, issues=issues)
