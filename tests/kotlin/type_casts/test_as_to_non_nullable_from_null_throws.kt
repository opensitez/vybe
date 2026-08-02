// vybe-test: kotlin/type_casts/test_as_to_non_nullable_from_null_throws
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            try {
                val value: String? = null
                val forced = value as String
                println("bad")
            } catch (e: Exception) {
                println("caught")
            }
        }

