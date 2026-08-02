// vybe-test: kotlin/kotlin_annotations_runtime/test_annotation_retrieval_after_instance_creation
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val service = Service()
            val annClass = service::class.java.getAnnotation(Marker::class.java)
            val annName = service::class.java.name
            __check((annClass?.kind).toString(), "service")
            __check((annName.endsWith("Service")).toString(), "true")
        }
