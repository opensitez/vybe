// vybe-test: kotlin/kotlin_annotations_runtime/test_annotation_with_local_class
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            @Marker("local")
            class Local
            val ann = Local::class.java.getAnnotation(Marker::class.java)
            __check((ann?.kind).toString(), "local")
        }
