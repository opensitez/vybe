// vybe-test: kotlin/inner_classes/test_inner_class_accesses_outer_mutable_state
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Store {
            var value = 1
            inner class Bump {
                fun add(v: Int) {
                    value += v
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
            val store = Store()
            val bump = store.Bump()
            bump.add(5)
            __check((store.value).toString(), "6")
        }
