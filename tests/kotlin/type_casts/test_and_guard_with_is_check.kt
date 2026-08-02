// vybe-test: kotlin/type_casts/test_and_guard_with_is_check
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val item: Any = "abc"
            if (item is String && item.isNotEmpty()) {
                println(item.length)
            } else {
                println(0)
            }
        }

