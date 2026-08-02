// vybe-test: kotlin/type_casts/test_as_array_to_incompatible_component_type_fails
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val value: Any = arrayOf(1, 2, 3)
            try {
                val casted = value as Array<String>
                println(casted[0])
            } catch (e: Exception) {
                println("caught")
            }
        }

