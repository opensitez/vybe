// vybe-test: kotlin/visibility/test_private_class_is_file_local_access_control
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

private class Hidden {
            val value = 8
        }

        fun spawn(): Int {
            return Hidden().value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((spawn()).toString(), "8")
        }
