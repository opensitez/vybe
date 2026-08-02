// vybe-test: kotlin/companion_objects/test_companion_object_companion_init_block_runs_once
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Probe {
            companion object {
                var value = 0
                init {
                    value = 7
                }
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Probe.value).toString(), "7")
            __check((Probe.value).toString(), "7")
        }
