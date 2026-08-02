// vybe-test: kotlin/kotlin_annotations_runtime/test_annotation_array_inheritance_not_automatic
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

@Marker("base")
        open class Base

        class Child : Base()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = Base::class.java.getAnnotation(Marker::class.java)
            val child = Child::class.java.getAnnotation(Marker::class.java)
            __check((base?.kind).toString(), "base")
            __check((child == null).toString(), "true")
        }
