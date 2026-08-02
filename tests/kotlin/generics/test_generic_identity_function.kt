// vybe-test: kotlin/generics/test_generic_identity_function
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> identity(value: T): T {
            return value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((identity(3)).toString(), "3")
            __check((identity("ok")).toString(), "ok")
        }
