// vybe-test: kotlin/companion_objects/test_companion_object_with_init_block_runs_once
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Loader {
            companion object {
                var status = "cold"

                init {
                    status = "warm"
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
            __check((Loader.status).toString(), "warm")
            __check((Loader.status).toString(), "warm")
        }
