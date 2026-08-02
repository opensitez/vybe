// vybe-test: kotlin/evaluation_order/test_for_each_loop_evaluates_initializer_before_body
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun main() {
            var order = ""
            val list = run {
                order += "init"
                listOf(1, 2)
            }
            var sum = 0
            for (x in list) {
                order += "-"
                sum += x
            }
            println(sum)
            println(order)
        }

