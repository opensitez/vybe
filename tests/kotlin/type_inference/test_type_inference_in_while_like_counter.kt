// vybe-test: kotlin/type_inference/test_type_inference_in_while_like_counter
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun main() {
            var i = 0
            var total = 0
            while (i < 3) {
                total += i
                i++
            }
            println(total)
        }

