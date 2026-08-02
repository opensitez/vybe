// vybe-test: kotlin/kotlin_string_interpolation_advanced/test_interpolation_in_loop_accumulation
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_interpolation_advanced.rs

fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println("sum=${'$'}total")
        }

