// vybe-test: kotlin/control_flow/test_until_range_excludes_upper_bound
// origin: languages/kotlin/tests/kotlin/test_control_flow.rs

fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }

