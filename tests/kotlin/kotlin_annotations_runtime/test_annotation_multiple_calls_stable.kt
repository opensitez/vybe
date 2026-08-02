// vybe-test: kotlin/kotlin_annotations_runtime/test_annotation_multiple_calls_stable
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val first = Service::class.java.getAnnotation(Marker::class.java)
            val second = Service::class.java.getAnnotation(Marker::class.java)
            __check((first == second).toString(), "true")
            __check((first?.kind).toString(), "service")
        }
