// vybe-test: kotlin/kotlin_uuid/test_uuid_random_uuid_format_shape
// origin: languages/kotlin/tests/kotlin/test_kotlin_uuid.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val id = java.util.UUID.randomUUID()
            val text = id.toString()
            __check((text.length).toString(), "36")
            __check((text[8] == '-').toString(), "true")
            __check((text[13] == '-').toString(), "true")
            __check((text[18] == '-').toString(), "true")
            __check((text[23] == '-').toString(), "true")
        }
