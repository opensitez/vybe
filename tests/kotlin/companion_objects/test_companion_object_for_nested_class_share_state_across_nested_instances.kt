// vybe-test: kotlin/companion_objects/test_companion_object_for_nested_class_share_state_across_nested_instances
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Container {
            class Unit {
                companion object {
                    var count = 0
                    fun use(): Int {
                        count += 1
                        return count
                    }
                }
            }

            fun call(): Int = Unit.use()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = Container.Unit.use()
            val two = Container.Unit()
            val three = two.call()
            val four = Container.Unit.use()
            __check((one).toString(), "1")
            __check((three).toString(), "2")
            __check((four).toString(), "3")
        }
