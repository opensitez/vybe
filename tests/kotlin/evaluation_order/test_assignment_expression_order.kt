// vybe-test: kotlin/evaluation_order/test_assignment_expression_order
// origin: languages/kotlin/tests/kotlin/test_evaluation_order.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var a = 0
            var b = 0
            val list = mutableListOf<Int>()
            fun left() { a = 1
list.add(1) }
            fun right() { b = 2
list.add(2) }
            left()
            right()
            __check((a + b).toString(), "3")
            __check((list.joinToString(",")).toString(), "1,2")
        }
