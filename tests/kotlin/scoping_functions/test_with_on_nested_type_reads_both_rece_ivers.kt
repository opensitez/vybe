// vybe-test: kotlin/scoping_functions/test_with_on_nested_type_reads_both_rece_ivers
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Holder {
            var value = "h"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = mutableListOf("x", "y")
            val result = with(list) {
                with(Holder()) {
                    list.add("z")
                    value = list.first()
                    value
                }
            }
            __check((result).toString(), "x")
            __check((list.size).toString(), "3")
        }
