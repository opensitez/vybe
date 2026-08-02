// vybe-test: kotlin/spread_arguments/test_spread_vararg_reference_within_class
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

class Combiner {
            fun build(prefix: String, vararg values: Int): String {
                return prefix + values.joinToString(",")
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Combiner()
            val a = intArrayOf(7, 8)
            __check((c.build("n", *a)).toString(), "n7,8")
        }
