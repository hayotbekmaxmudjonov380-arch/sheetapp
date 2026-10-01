"""Formula dvigateli: =A1+B1, =SUM(A1:A10), ='Varaq'!B2, SUMIF/SUMIFS, IF..."""
import re
from datetime import date, datetime, time

from .storage import col_letter, letter_col

ERR_CYCLE = "#CYCLE!"
ERR_REF = "#REF!"
ERR_NAME = "#NAME?"
ERR_VALUE = "#VALUE!"
ERR_DIV = "#DIV/0!"
ERR_PARSE = "#PARSE!"

_REF_RE = re.compile(
    r"^(?:(?:'(?P<q>[^']+)'|(?P<i>[A-Za-z_\u0400-\u04FF][\w\u0400-\u04FF]*))!)?"
    r"\$?(?P<c1>[A-Za-z]{1,3})\$?(?P<r1>\d+)"
    r"(?::\$?(?P<c2>[A-Za-z]{1,3})\$?(?P<r2>\d+))?$",
    re.UNICODE)

_TOKEN_RE = re.compile(r"""
    (?P<str>"(?:[^"]|"")*")
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<func>[A-Za-z_\u0400-\u04FF][\w\u0400-\u04FF]*(?=\s*\())
  | (?P<ref>(?:'[^']+'|[A-Za-z_\u0400-\u04FF][\w\u0400-\u04FF]*)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?
           |\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
  | (?P<op><>|<=|>=|[-+*/^&=<>])
  | (?P<lp>\()
  | (?P<rp>\))
  | (?P<comma>,)
  | (?P<ws>\s+)
""", re.VERBOSE | re.UNICODE)


class FormulaError(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.code = code


def tokenize(s):
    toks, i = [], 0
    while i < len(s):
        m = _TOKEN_RE.match(s, i)
        if not m:
            raise FormulaError(ERR_PARSE)
        i = m.end()
        kind = m.lastgroup
        if kind == "ws":
            continue
        toks.append((kind, m.group()))
    return toks


# ---------- parser ----------
class Parser:
    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def next(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expect(self, kind):
        k, v = self.next()
        if k != kind:
            raise FormulaError(ERR_PARSE)
        return v

    def parse(self):
        node = self.cmp()
        if self.i != len(self.t):
            raise FormulaError(ERR_PARSE)
        return node

    def cmp(self):
        n = self.concat()
        while self.peek()[0] == "op" and self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.next()[1]
            n = ("bin", op, n, self.concat())
        return n

    def concat(self):
        n = self.add()
        while self.peek() == ("op", "&"):
            self.next()
            n = ("bin", "&", n, self.add())
        return n

    def add(self):
        n = self.mul()
        while self.peek()[0] == "op" and self.peek()[1] in ("+", "-"):
            op = self.next()[1]
            n = ("bin", op, n, self.mul())
        return n

    def mul(self):
        n = self.pow()
        while self.peek()[0] == "op" and self.peek()[1] in ("*", "/"):
            op = self.next()[1]
            n = ("bin", op, n, self.pow())
        return n

    def pow(self):
        n = self.unary()
        while self.peek() == ("op", "^"):
            self.next()
            n = ("bin", "^", n, self.unary())
        return n

    def unary(self):
        if self.peek()[0] == "op" and self.peek()[1] in ("-", "+"):
            op = self.next()[1]
            return ("un", op, self.unary())
        return self.primary()

    def primary(self):
        k, v = self.next()
        if k == "num":
            return ("num", float(v))
        if k == "str":
            return ("str", v[1:-1].replace('""', '"'))
        if k == "func":
            self.expect("lp")
            args = []
            if self.peek()[0] != "rp":
                args.append(self.cmp())
                while self.peek()[0] == "comma":
                    self.next()
                    args.append(self.cmp())
            self.expect("rp")
            return ("call", v.upper(), args)
        if k == "ref":
            return ("ref", v)
        if k == "lp":
            n = self.cmp()
            self.expect("rp")
            return n
        raise FormulaError(ERR_PARSE)


def parse_formula(f):
    return Parser(tokenize(f[1:] if f.startswith("=") else f)).parse()


# ---------- evaluator ----------
class Evaluator:
    def __init__(self, store, current_sheet=None):
        self.store = store
        self.cur = current_sheet
        self.memo = {}
        self.stack = set()
        self._trees = {}

    def invalidate(self):
        self.memo.clear()
        self.stack.clear()

    # ---- cell qiymati ----
    def raw_value(self, sid, r, c):
        raw = self.store.get_raw(sid, r, c)
        if raw is None or raw == "":
            return ""
        if isinstance(raw, str) and raw.startswith("="):
            return self.formula_value(sid, raw)
        return literal(raw)

    def formula_value(self, sid, raw):
        key = (sid, raw)
        if key in self.memo:
            return self.memo[key]
        if key in self.stack:
            raise FormulaError(ERR_CYCLE)
        try:
            tree = self._trees.get(raw)
            if tree is None:
                tree = parse_formula(raw)
                self._trees[raw] = tree
            self.stack.add(key)
            try:
                val = self.eval_node(tree, sid)
            finally:
                self.stack.discard(key)
        except FormulaError as e:
            val = e.code
        except Exception:
            val = ERR_PARSE
        self.memo[key] = val
        return val

    def ref_value(self, refstr, cur_sheet):
        m = _REF_RE.match(refstr)
        if not m:
            raise FormulaError(ERR_REF)
        sname = m.group("q") or m.group("i")
        sid = cur_sheet if sname is None else self.store.sheet_id(sname)
        if sid is None:
            raise FormulaError(ERR_REF)
        c1, r1 = letter_col(m.group("c1")), int(m.group("r1")) - 1
        if m.group("c2") is None:
            return self.raw_value(sid, r1, c1)
        c2, r2 = letter_col(m.group("c2")), int(m.group("r2")) - 1
        return self.range_values(sid, r1, c1, r2, c2)

    def range_values(self, sid, r1, c1, r2, c2):
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        return [("cell", sid, r, c)
                for r in range(r1, r2 + 1)
                for c in range(c1, c2 + 1)]

    def cell_of(self, item, cur_sheet):
        if item[0] == "cell":
            return item[1], item[2], item[3]
        if item[0] == "ref":
            m = _REF_RE.match(item[1])
            sname = m.group("q") or m.group("i")
            sid = cur_sheet if sname is None else self.store.sheet_id(sname)
            if sid is None:
                raise FormulaError(ERR_REF)
            return sid, int(m.group("r1")) - 1, letter_col(m.group("c1"))
        raise FormulaError(ERR_VALUE)

    def resolve(self, node, cur_sheet):
        """Argumentni qiymatga aylantiradi (range -> ro'yxat)."""
        if node[0] == "ref":
            m = _REF_RE.match(node[1])
            if m and m.group("c2") is not None:
                return self.ref_value(node[1], cur_sheet)
            return self.ref_value(node[1], cur_sheet)
        return self.eval_node(node, cur_sheet)

    def eval_node(self, n, sid):
        t = n[0]
        if t == "num":
            return n[1]
        if t == "str":
            return n[1]
        if t == "ref":
            return self.ref_value(n[1], sid)
        if t == "un":
            v = num(self.eval_node(n[2], sid))
            return -v if n[1] == "-" else v
        if t == "bin":
            return self.binop(n, sid)
        if t == "call":
            return self.call(n, sid)
        raise FormulaError(ERR_PARSE)

    def binop(self, n, sid):
        op = n[1]
        a = self.eval_node(n[2], sid)
        b = self.eval_node(n[3], sid)
        if isinstance(a, str) and a.startswith("#") and a in _ERRS:
            return a
        if isinstance(b, str) and b in _ERRS:
            return b
        if op == "&":
            return fmt(a) + fmt(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return compare(a, b, op)
        try:
            x, y = num(a), num(b)
        except Exception:
            raise FormulaError(ERR_VALUE)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise FormulaError(ERR_DIV)
            return x / y
        if op == "^":
            return x ** y
        raise FormulaError(ERR_PARSE)

    # ---- funksiyalar ----
    def call(self, n, sid):
        name, args = n[1], n[2]
        fn = _FUNCS.get(name)
        if fn is None:
            raise FormulaError(ERR_NAME)
        # RANGE arg shakllarini tayyorlash
        resolved = []
        for a in args:
            if a[0] == "ref" and _REF_RE.match(a[1]) and _REF_RE.match(a[1]).group("c2"):
                resolved.append(self.ref_value(a[1], sid))
            else:
                resolved.append(self.eval_node(a, sid))
        return fn(self, resolved, sid)

    def flat(self, vals):
        out = []
        for v in vals:
            if isinstance(v, list):
                out.extend(v)
            else:
                out.append(v)
        return out

    def numbers(self, vals):
        out = []
        for v in self.flat(vals):
            if isinstance(v, tuple) and v and v[0] == "cell":
                v = self.raw_value(v[1], v[2], v[3])
                if isinstance(v, tuple):
                    continue
            if isinstance(v, list):
                out.extend(self.numbers(v))
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append(float(v))
            elif isinstance(v, str) and v != "" and v not in _ERRS:
                try:
                    out.append(float(v))
                except ValueError:
                    pass
        return out


_ERRS = {ERR_CYCLE, ERR_REF, ERR_NAME, ERR_VALUE, ERR_DIV, ERR_PARSE}


def num(v):
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0
    if isinstance(v, str):
        if v in _ERRS:
            raise FormulaError(v)
        try:
            return float(v.replace(",", ""))
        except ValueError:
            raise FormulaError(ERR_VALUE)
    if isinstance(v, (date, datetime, time)):
        raise FormulaError(ERR_VALUE)
    raise FormulaError(ERR_VALUE)


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def compare(a, b, op):
    if isinstance(a, str) and a in _ERRS:
        return a
    if isinstance(b, str) and b in _ERRS:
        return b
    try:
        if isinstance(a, str) or isinstance(b, str):
            if isinstance(a, str) and isinstance(b, str):
                x, y = a.upper(), b.upper()
            else:
                x, y = str(a), str(b)
        else:
            x, y = num(a), num(b)
    except FormulaError as e:
        return e.code
    res = {"=": x == y, "<>": x != y,
           "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]
    return res


def literal(raw):
    if isinstance(raw, (int, float)):
        return raw
    s = str(raw).strip()
    if s == "":
        return ""
    try:
        if re.fullmatch(r"-?\d+", s):
            return int(s)
        if re.fullmatch(r"-?\d*\.\d+", s):
            return float(s)
    except ValueError:
        pass
    return s


def f_sum(e, a, sid):
    return sum(e.numbers(a)) or 0.0


def f_avg(e, a, sid):
    n = e.numbers(a)
    if not n:
        raise FormulaError(ERR_DIV)
    return sum(n) / len(n)


def f_min(e, a, sid):
    n = e.numbers(a)
    return min(n) if n else 0.0


def f_max(e, a, sid):
    n = e.numbers(a)
    return max(n) if n else 0.0


def f_count(e, a, sid):
    return float(len(e.numbers(a)))


def f_counta(e, a, sid):
    return float(sum(1 for v in e.flat(a) if v not in ("", None)))


def f_if(e, a, sid):
    cond = a[0]
    t = bool(cond) if not isinstance(cond, (int, float)) else cond != 0
    if isinstance(cond, str):
        t = cond.upper() == "TRUE"
    if t:
        return a[1] if len(a) > 1 else True
    return a[2] if len(a) > 2 else False


def f_round(e, a, sid):
    d = int(num(a[1])) if len(a) > 1 else 0
    return round(num(a[0]), d)


def f_abs(e, a, sid):
    return abs(num(a[0]))


def f_and(e, a, sid):
    return all(bool(v) for v in e.flat(a))


def f_or(e, a, sid):
    return any(bool(v) for v in e.flat(a))


def f_not(e, a, sid):
    return not bool(a[0])


def f_len(e, a, sid):
    return float(len(fmt(a[0])))


def f_concat(e, a, sid):
    return "".join(fmt(v) for v in e.flat(a))


def _match(cell_val, crit):
    """SUMIF/SUMIFS mezonini baholaydi."""
    if isinstance(crit, str):
        m = re.match(r"^(<=|>=|<>|<|>)(.*)$", crit)
        if m:
            op, rhs = m.group(1), m.group(2)
            try:
                return compare(cell_val, float(rhs), op)
            except (ValueError, FormulaError):
                return compare(cell_val, rhs, op)
        if crit.endswith("*") or crit.startswith("*"):
            s = fmt(cell_val).upper()
            return re.fullmatch(crit.replace("*", ".*").upper(), s) is not None
        return fmt(cell_val).upper() == crit.upper()
    return compare(cell_val, crit, "=")


def f_sumif(e, a, sid):
    rng, crit = a[0], a[1]
    total = a[2] if len(a) > 2 else rng
    if not isinstance(rng, list) or not isinstance(total, list):
        raise FormulaError(ERR_VALUE)
    s = 0.0
    for i, cell in enumerate(rng):
        cv = e.raw_value(*cell[1:]) if cell[0] == "cell" else cell
        if _match(cv, crit):
            if i < len(total):
                t = e.raw_value(*total[i][1:]) if total[i][0] == "cell" else total[i]
                try:
                    s += num(t)
                except FormulaError:
                    pass
    return s


def f_sumifs(e, a, sid):
    total = a[0]
    pairs = a[1:]
    if len(pairs) % 2:
        raise FormulaError(ERR_VALUE)
    s = 0.0
    for i, cell in enumerate(total):
        ok = True
        for j in range(0, len(pairs), 2):
            rng, crit = pairs[j], pairs[j + 1]
            if i >= len(rng):
                ok = False
                break
            cv = e.raw_value(*rng[i][1:]) if rng[i][0] == "cell" else rng[i]
            if not _match(cv, crit):
                ok = False
                break
        if ok:
            t = e.raw_value(*cell[1:]) if cell[0] == "cell" else cell
            try:
                s += num(t)
            except FormulaError:
                pass
    return s


_FUNCS = {
    "SUM": f_sum, "AVERAGE": f_avg, "MIN": f_min, "MAX": f_max,
    "COUNT": f_count, "COUNTA": f_counta, "IF": f_if, "ROUND": f_round,
    "ABS": f_abs, "AND": f_and, "OR": f_or, "NOT": f_not,
    "LEN": f_len, "CONCAT": f_concat, "CONCATENATE": f_concat,
    "SUMIF": f_sumif, "SUMIFS": f_sumifs,
}
