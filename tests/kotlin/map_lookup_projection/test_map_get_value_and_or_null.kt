// vybe-test: kotlin/map_lookup_projection/test_map_get_value_and_or_null
// origin: languages/kotlin/tests/kotlin/test_map_lookup_projection.rs

fun main() {
            val source = mapOf("a" to 1)
            try {
                println(source.getValue("a"))
                println(source.getValue("b"))
            } catch (e: Exception) {
                println("err")
            }
        }

