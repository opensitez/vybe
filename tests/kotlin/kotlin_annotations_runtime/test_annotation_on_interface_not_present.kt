// vybe-test: kotlin/kotlin_annotations_runtime/test_annotation_on_interface_not_present
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

interface SampleInterface

        class Impl : SampleInterface

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ann = SampleInterface::class.java.getAnnotation(Marker::class.java)
            val implAnn = Impl::class.java.getAnnotation(Marker::class.java)
            __check((ann == null).toString(), "true")
            __check((implAnn == null).toString(), "true")
        }
