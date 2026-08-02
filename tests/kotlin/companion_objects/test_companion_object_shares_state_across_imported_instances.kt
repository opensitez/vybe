// vybe-test: kotlin/companion_objects/test_companion_object_shares_state_across_imported_instances
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Registry {
            companion object {
                var values = 0
            }
        }

        fun bump() {
            Registry.values += 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Registry.values).toString(), "0")
            bump()
            bump()
            __check((Registry.values).toString(), "2")
        }
