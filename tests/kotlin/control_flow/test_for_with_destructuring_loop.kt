// vybe-test: kotlin/control_flow/test_for_with_destructuring_loop
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var total = 0
            for ((x, y) in arrayOf(Pair(1, 2), Pair(3, 4))) {
                total += x + y
            }
            println(total)
        }

