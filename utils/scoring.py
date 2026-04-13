def clamp(value, min_value, max_value):
    return max(min_value, min(value, max_value))


def piecewise_score(x, xs, ys):
    """
    x: 실제 달성도(%)
    xs: 기준 구간값 리스트 (내림차순 또는 오름차순 가능)
    ys: 각 구간 점수 리스트
    구간 사이 값은 선형보간
    """
    if len(xs) != len(ys):
        raise ValueError("xs와 ys 길이가 다릅니다.")

    pairs = sorted(zip(xs, ys), key=lambda t: t[0])
    xs_sorted = [p[0] for p in pairs]
    ys_sorted = [p[1] for p in pairs]

    if x <= xs_sorted[0]:
        return ys_sorted[0]
    if x >= xs_sorted[-1]:
        return ys_sorted[-1]

    for i in range(len(xs_sorted) - 1):
        x1, x2 = xs_sorted[i], xs_sorted[i + 1]
        y1, y2 = ys_sorted[i], ys_sorted[i + 1]
        if x1 <= x <= x2:
            if x2 == x1:
                return y1
            ratio = (x - x1) / (x2 - x1)
            return y1 + (y2 - y1) * ratio

    return ys_sorted[-1]


def size_factor(actual, group_avg, min_v=0.9, max_v=1.1):
    if group_avg <= 0:
        return 1.0
    factor = 0.7 + 0.3 * (actual / group_avg)
    return clamp(factor, min_v, max_v)


def bonus_by_step(rate, start_rate, step_rate, step_point, max_bonus):
    """
    예: 105% 초과 5%당 0.2점, 최대 2점
    """
    if rate <= start_rate:
        return 0.0
    steps = int((rate - start_rate) // step_rate)
    bonus = steps * step_point
    return min(bonus, max_bonus)