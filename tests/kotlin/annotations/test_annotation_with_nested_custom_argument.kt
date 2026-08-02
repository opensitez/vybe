// vybe-test: kotlin/annotations/test_annotation_with_nested_custom_argument
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

annotation class Tag(val value: String)
        annotation class Marker(val tag: Tag)

        @Marker(Tag("critical"))
        class Item

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Item().javaClass.simpleName).toString(), "Item")
        }
