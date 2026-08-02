// vybe-test: kotlin/collections/test_array_zero_initialization_with_repeat
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun main() {
            val zeros = Array(4) { 0 }
            println(zeros.size)
            var total = 0
            for (value in zeros) {
                total += value
            }
            println(total)
        }

