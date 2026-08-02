// vybe-test: kotlin/annotations/test_annotation_on_constructor_parameter
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class User(
            @Deprecated("unused") val id: Int,
            @Suppress("UNUSED_PARAMETER") val name: String
        )

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val user = User(1, "alice")
            __check((user.id).toString(), "1")
            __check((user.name).toString(), "alice")
        }
