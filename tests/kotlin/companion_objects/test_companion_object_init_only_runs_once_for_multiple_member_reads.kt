// vybe-test: kotlin/companion_objects/test_companion_object_init_only_runs_once_for_multiple_member_reads
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

var init_log = ""

        class Tracker {
            companion object {
                init {
                    init_log += "init;"
                }

                val tag = "ok"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((init_log).toString(), "")
            __check((Tracker.tag).toString(), "ok")
            __check((Tracker.tag).toString(), "ok")
            __check((init_log).toString(), "init;")
        }
