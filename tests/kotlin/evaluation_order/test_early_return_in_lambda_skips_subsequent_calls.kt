// vybe-test: kotlin/evaluation_order/test_early_return_in_lambda_skips_subsequent_calls
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun main() {
            var order = ""
            fun side(): Int {
                order += "s"
                return 2
            }
            val out = run {
                order += "r"
                1
            }
            if (out > 0) {
                println(side())
            } else {
                println(0)
            }
            println(order)
        }

