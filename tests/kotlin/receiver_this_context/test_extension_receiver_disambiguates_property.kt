// vybe-test: kotlin/receiver_this_context/test_extension_receiver_disambiguates_property
// origin: languages/kotlin/tests/kotlin/test_receiver_this_context.rs

class Context {
            val label = "root"
            inner class Node {
                val label = "node"
                fun describe(): String = "${'$'}{this@Context.label}/${'$'}{label}"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Context().Node().describe()).toString(), "root/node")
        }
