// vybe-test: kotlin/type_casts/test_casting_nullable_to_non_nullable_is_forced_and_throws
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun main() {
            val source: Any? = null
            val direct: String? = source as String?
            println(direct == null)

            try {
                val strict: String = source as String
                println(strict)
            } catch (e: Exception) {
                println("caught")
            }
        }

