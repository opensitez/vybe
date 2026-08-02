// vybe-test: kotlin/spread_arguments/test_spread_with_object_array
// origin: languages/kotlin/tests/kotlin/test_spread_arguments.rs

fun describe(prefix: String, vararg values: Any): String {
            return prefix + values.size
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val vals: Array<Any> = arrayOf(1, "x", true)
            __check((describe("n", *vals)).toString(), "n3")
        }
