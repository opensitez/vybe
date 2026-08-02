// vybe-test: kotlin/data_class_destructuring/test_destructure_map_entries
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

fun main() {
            val values = mapOf("a" to 1, "b" to 2)
            var total = 0
            for ((key, value) in values) {
                if (key == "a") {
                    total += value
                }
            }
            println(total)
        }

