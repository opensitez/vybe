// vybe-test: kotlin/initialization_order/test_initialization_of_local_class_occurs_on_use
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            class Local {
                init { __check(("local").toString(), "start") }
            }

            __check(("start").toString(), "local")
            Local()
            __check(("end").toString(), "end")
        }
