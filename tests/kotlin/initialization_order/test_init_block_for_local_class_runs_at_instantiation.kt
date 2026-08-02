// vybe-test: kotlin/initialization_order/test_init_block_for_local_class_runs_at_instantiation
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local {
                init { __check(("local-init").toString(), "a") }
            }

            __check(("a").toString(), "local-init")
            Local()
            __check(("b").toString(), "b")
        }
