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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val expr = Expr.Add(Expr.Value(3), Expr.Negate(Expr.Value(2)))
            __check((evaluate(expr)).toString(), "1")
        }
