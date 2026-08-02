// vybe-test: kotlin/kotlin_progressions/test_int_progression_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun main() {
            val values = 1..5
            var total = 0
            for (v in values) { total += v }
            println(total)
            println(values.step)
            println(values.last)
        }

