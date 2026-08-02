// vybe-test: kotlin/kotlin_annotations_runtime/test_runtime_parameter_annotation_retrieved
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val method = AnnotatedModel::class.java.getDeclaredMethod("tagged", Int::class.java)
            val param = method.parameters[0]
            val ann = param.getAnnotation(Marker::class.java)
            __check((ann?.kind).toString(), "param")
            __check((param.type.name).toString(), "int")
        }
