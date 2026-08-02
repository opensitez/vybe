// vybe-test: kotlin/visibility/test_inner_class_can_read_outer_private_member
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Vault {
            private val secret = "vault"

            inner class Reader {
                fun open(): String = secret
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val vault = Vault()
            __check((vault.Reader().open()).toString(), "vault")
        }
