// vybe-test: kotlin/kotlin_annotations_runtime/test_runtime_method_annotation_retrieved
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val method = AnnotatedModel::class.java.getDeclaredMethod("tagged", Int::class.java)
            val ann = method.getAnnotation(Marker::class.java)
            __check((ann?.kind).toString(), "ctor")
            __check((method.name).toString(), "tagged")
        }
