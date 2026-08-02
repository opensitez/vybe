// vybe-test: kotlin/mutable_list_apis/test_mutable_list_index_iteration
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun main() {
            val values = mutableListOf(1, 2, 3)
            var acc = 0
            for ((index, value) in values.withIndex()) {
                acc += index + value
            }
            println(acc)
        }

