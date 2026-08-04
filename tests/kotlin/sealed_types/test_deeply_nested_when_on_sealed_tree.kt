// vybe-test: kotlin/sealed_types/test_deeply_nested_when_on_sealed_tree
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Expr {
            class Value(val value: Int) : Expr()
            class Add(val left: Expr, val right: Expr) : Expr()
            class Negate(val source: Expr) : Expr()
        }

        fun evaluate(expr: Expr): Int {
            return when (expr) {
                is Expr.Value -> expr.value
                is Expr.Negate -> -evaluate(expr.source)
                is Expr.Add -> evaluate(expr.left) + evaluate(expr.right)
            }
        }

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val expr = Expr.Add(Expr.Value(3), Expr.Negate(Expr.Value(2)))
            __p((evaluate(expr)).toString())
        
__check("1")
}
