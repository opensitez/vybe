// vybe-test: kotlin/when_expressions/test_when_subject_capture_with_is_and_casts
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

interface Shape { fun kind(): String }
        class Dot : Shape {
            override fun kind(): String = "dot"
        }
        class Box(val size: Int) : Shape {
            override fun kind(): String = "box:" + size
        }

        fun describe(shape: Shape): String {
            return when (shape) {
                is Dot -> "dot"
                is Box -> "box=" + shape.size
                else -> "unknown"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(Dot())).toString(), "dot")
            __check((describe(Box(7))).toString(), "box=7")
        }
