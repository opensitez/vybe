// vybe-test: kotlin/type_inference/test_type_inference_for_for_loop_items
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun main() {
            var sum = 0
            for (v in listOf(1, 2, 3)) {
                sum += v
            }
            println(sum)
        }

