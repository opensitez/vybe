// vybe-test: kotlin/initialization_order/test_local_class_initialization_runs_after_function_code_before_use
// origin: languages/kotlin/tests/kotlin/test_initialization_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("pre").toString(), "pre")

            class Holder {
                val value = "init"

                init {
                    __check((value).toString(), "post")
                }
            }

            __check(("post").toString(), "init")
            Holder()
            __check(("done").toString(), "done")
        }
