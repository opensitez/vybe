// vybe-test: kotlin/lateinit_properties/test_lateinit_with_try_after_set
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Probe {
            lateinit var text: String
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Probe()
            p.text = "x"
            val result = try {
                p.text
                "after-set"
            } catch (e: Exception) {
                "bad"
            }
            __check((result).toString(), "after-set")
        }
