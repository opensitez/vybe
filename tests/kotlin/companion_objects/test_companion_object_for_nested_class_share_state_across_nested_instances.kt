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

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val one = Container.Unit.use()
            val two = Container.Unit()
            val three = two.call()
            val four = Container.Unit.use()
            __p((one).toString())
            __p((three).toString())
            __p((four).toString())
        
__check("1\n2\n3")
}
