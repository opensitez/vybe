// vybe-test: kotlin/kotlin_array_basic_ops/test_array_iteration_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_array_basic_ops.rs

fun main() {
            val a = arrayOf(2, 4, 6)
            var acc = 0
            for (v in a) {
                acc += v
            }
            println(acc)
        }

