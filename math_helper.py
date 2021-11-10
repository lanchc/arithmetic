# -*- coding:utf-8 _*-
# @time     2021/11/10 1:52 下午
# @explain  arithmetic
# 基础常用方法
import random


# 计算结果
def compute_outcome(a, b, operator):
    if operator == '-':
        return a - b
    else:
        return a + b


# 计算结果 - 混合运算
def compute_outcome_mingle(a, b, c, sy_a, ay_b):
    temp_r_a = compute_outcome(a, b, sy_a)
    temp_r_b = compute_outcome(temp_r_a, c, ay_b)
    return temp_r_b


# 格式化输出模版
def format_basics_text(a, b, c, d):
    style_rd = random.randint(1, 3)
    if 1 == style_rd:
        return '{0: >2} {1:} {2: >2} = （   ）'.format(int(a), str(b), int(c), int(d))
    elif 2 == style_rd:
        return '（   ） {1:} {2: >2} = {3: >2}'.format(int(a), str(b), int(c), int(d))
    elif 3 == style_rd:
        return '{0: >2} {1:} （   ）= {3: >2}'.format(int(a), str(b), int(c), int(d))
    else:
        return '{0: >2} {1:} {2: >2} = {3: >2}'.format(int(a), str(b), int(c), int(d))


