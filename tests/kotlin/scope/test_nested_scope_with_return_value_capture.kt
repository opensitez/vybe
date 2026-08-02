// vybe-test: kotlin/scope/test_nested_scope_with_return_value_capture
// origin: languages/kotlin/tests/kotlin/test_scope.rs

fun make() : Int {
            val prefix = 1
            fun inner(): Int {
                val suffix = 2
                return prefix + suffix
            }
            return inner()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((make()).toString(), "3")
        }
