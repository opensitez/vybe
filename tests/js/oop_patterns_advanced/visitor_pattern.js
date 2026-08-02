// vybe-test: js/oop_patterns_advanced/visitor_pattern
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

class NumberExpr { constructor(v) { this.v=v; } accept(visitor) { return visitor.visitNumber(this); } }
class AddExpr { constructor(l,r) { this.l=l; this.r=r; } accept(visitor) { return visitor.visitAdd(this); } }
class Evaluator {
    visitNumber(e) { return e.v; }
    visitAdd(e) { return e.l.accept(this) + e.r.accept(this); }
}
const expr = new AddExpr(new NumberExpr(3), new AddExpr(new NumberExpr(4), new NumberExpr(5)));
__check(__line(expr.accept(new Evaluator())), "12");
