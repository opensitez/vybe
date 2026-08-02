// vybe-test: kotlin/kotlin_this_qualifiers/test_this_parameter_capture
// origin: languages/kotlin/tests/kotlin/test_kotlin_this_qualifiers.rs

class Chain {
            val value = 5

            fun mark(prefix: String): String {
                return this.toStringPrefix(prefix)
            }

            fun toStringPrefix(prefix: String): String {
                return prefix + this.value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Chain().mark("v=")).toString(), "v=5")
        }
