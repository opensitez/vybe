// vybe-test: kotlin/kotlin_uuid/test_uuid_parse_invalid_throws
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun main() {
            try {
                java.util.UUID.fromString("not-a-uuid")
                println("bad")
            } catch (e: IllegalArgumentException) {
                println(e::class.simpleName)
            }
        }

