// vybe-test: kotlin/scope/test_when_branch_scoped_names
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun describe(value: Int): String {
            return when (value) {
                1 -> {
                    val label = "one"
                    label
                }
                2 -> {
                    val label = "two"
                    label
                }
                else -> {
                    val label = "other"
                    label
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe(1)).toString(), "one")
            __check((describe(2)).toString(), "two")
            __check((describe(4)).toString(), "other")
        }
