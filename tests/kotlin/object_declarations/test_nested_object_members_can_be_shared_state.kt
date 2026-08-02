// vybe-test: kotlin/object_declarations/test_nested_object_members_can_be_shared_state
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

class Container {
            object State {
                var value = 0
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Container.State.value += 4
            Container.State.value += 1
            __check((Container.State.value).toString(), "5")
        }
