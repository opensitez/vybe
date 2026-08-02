// vybe-test: kotlin/data_classes/test_data_class_string_projection_is_stable
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Tag(val name: String, val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Tag("x", 1)
            val b = Tag("y", 1)
            val list = listOf(a, b)
            __check((list.joinToString("|") { it.toString() }).toString(), "Tag(name=x, value=1)|Tag(name=y, value=1)")
        }
