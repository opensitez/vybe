// vybe-test: kotlin/annotations/test_annotation_on_typealias_value
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

@Suppress("UNUSED")
@Deprecated("legacy")
class Score(val value: Int = 13)

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val value = Score(13)
    __check((value.value).toString(), "13")
}
