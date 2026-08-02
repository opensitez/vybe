// vybe-test: kotlin/initialization_order/test_derived_class_init_runs_after_base_init
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

open class Base {
            init {
                __check(("base").toString(), "base")
            }
        }

        class Child : Base() {
            init {
                __check(("child-1").toString(), "child-1")
            }

            init {
                __check(("child-2").toString(), "child-2")
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            Child()
        }
