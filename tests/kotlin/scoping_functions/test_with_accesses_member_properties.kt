// vybe-test: kotlin/scoping_functions/test_with_accesses_member_properties
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Box {
            val prefix = "left"
            val suffix = "right"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = with(Box()) {
                prefix + ":" + suffix
            }
            __check((text).toString(), "left:right")
        }
