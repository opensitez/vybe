// vybe-test: kotlin/annotations/test_annotation_on_class_constructor_parameter
// origin: languages/kotlin/tests/kotlin/test_annotations.rs

class Packet(@Deprecated("old") val id: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val p = Packet(9)
__check((p.id).toString(), "9") }
