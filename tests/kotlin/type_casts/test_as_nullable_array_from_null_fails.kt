// vybe-test: kotlin/type_casts/test_as_nullable_array_from_null_fails
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            try {
                val value: Any? = null
                val casted = value as Array<Int>
                println(casted.size)
            } catch (e: Exception) {
                println("caught")
            }
        }

