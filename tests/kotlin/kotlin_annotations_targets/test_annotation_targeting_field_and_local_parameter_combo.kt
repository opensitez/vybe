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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Payload("x").copyMarker).toString(), "x")
        }
