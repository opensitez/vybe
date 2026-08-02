// vybe-test: kotlin/labeled_control_flow/test_label_on_when_like_control_is_invalid
// origin: languages/kotlin/tests/kotlin/test_labeled_control_flow.rs

fun main() {
            var count = 0
            search@ for (n in listOf(1, 2, 3, 4)) {
                val labelValue = if (n == 3) {
                    continue@search
                } else if (n == 4) {
                    break@search
                } else {
                    n
                }
                count += labelValue
            }
            println(count)
        }

