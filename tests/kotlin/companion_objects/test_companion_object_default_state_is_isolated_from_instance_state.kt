// vybe-test: kotlin/companion_objects/test_companion_object_default_state_is_isolated_from_instance_state
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder {
            companion object {
                var global = 0
            }

            var local = 0

            init {
                local += 1
                global += local
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Holder()
            val second = Holder()
            __check((first.local).toString(), "1")
            __check((second.local).toString(), "1")
            __check((Holder.global).toString(), "2")
        }
