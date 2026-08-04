// vybe-test: kotlin/kotlin_annotations_targets/test_annotation_targeting_field_and_local_parameter_combo
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_targets.rs

@Target(AnnotationTarget.FIELD)
        annotation class FieldMark

        @Target(AnnotationTarget.VALUE_PARAMETER)
        annotation class FieldArg

        class Payload(@FieldArg val marker: String) {
            @FieldMark
            val copyMarker: String = marker
        }

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((Payload("x").copyMarker).toString())
        
__check("x")
}
