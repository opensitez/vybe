// vybe-test: kotlin/mutable_map_apis/test_mutable_map_get_value_throws_if_missing
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun main() {
            val values = mutableMapOf("a" to 1)
            try {
                println(values.getValue("b"))
            } catch (e: NoSuchElementException) {
                println("missing")
            }
        }

