// vybe-test: kotlin/named_arguments/test_named_arguments_boolean_flags
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun pack(includeA: Boolean = true, includeB: Boolean = false): String {
            return (if (includeA) "A" else "") + (if (includeB) "B" else "")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pack()).toString(), "A")
            __check((pack(includeB = true)).toString(), "AB")
            __check((pack(includeA = false, includeB = true)).toString(), "B")
        }
