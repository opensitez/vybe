// vybe-test: kotlin/receiver_this_context/test_this_label_in_extension_lambda
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Host {
            fun transform(): String {
                val f: Host.() -> String = {
                    "${'$'}{this::class.simpleName}"
                }
                return f()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Host().transform()).toString(), "Host")
        }
