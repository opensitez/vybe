// vybe-test: kotlin/initialization_order/test_init_blocks_execute_in_chain_before_instance_values_are_printed
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class LevelOne {
            open val base = "one"
            init {
                __p((base).toString())
            }
        }

        open class LevelTwo : LevelOne() {
            init {
                __p((base + "-two").toString())
            }
        }

        class LevelThree : LevelTwo() {
            init {
                __p((base + "-three").toString())
            }
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
            LevelThree()
        
__check("one\none-two\none-three")
}
