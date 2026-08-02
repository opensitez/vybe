// vybe-test: kotlin/named_arguments/test_named_arguments_uses_default_before_named_override
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun concat(a: String = "1", b: String, c: String = "3"): String {
            return a + b + c
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((concat(b = "2")).toString(), "123")
            __check((concat(a = "A", b = "2", c = "C")).toString(), "A2C")
        }
