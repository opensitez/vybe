// vybe-test: kotlin/scoping_functions/test_run_reads_and_transforms_receiver
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder {
            var value = 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = Holder().run {
                value *= 3
                value + 1
            }
            __check((result).toString(), "7")
        }
