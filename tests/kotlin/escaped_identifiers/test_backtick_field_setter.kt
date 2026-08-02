// vybe-test: kotlin/escaped_identifiers/test_backtick_field_setter
// origin: languages/kotlin/tests/kotlin/test_escaped_identifiers.rs

class Holder { var `mutable field` = 1
            fun inc() { `mutable field` += 2 }
        }
    fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val h = Holder()
h.inc()
__check((h.`mutable field`).toString(), "3") }
