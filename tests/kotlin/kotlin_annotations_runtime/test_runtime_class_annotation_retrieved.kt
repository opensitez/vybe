// vybe-test: kotlin/kotlin_annotations_runtime/test_runtime_class_annotation_retrieved
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ann = Service::class.java.getAnnotation(Marker::class.java)
            __check((ann?.kind).toString(), "service")
            __check((ann != null).toString(), "true")
        }
