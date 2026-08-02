// vybe-test: kotlin/kotlin_annotations_runtime/test_annotations_apply_to_multiple_targets
// origin: languages/kotlin/tests/kotlin/test_kotlin_annotations_runtime.rs

class User {
            @Marker("value")
            var value = 0

            @Marker("action")
            fun action() {}
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val field = User::class.java.getDeclaredField("value")
            val method = User::class.java.getDeclaredMethod("action")
            val a1 = field.getAnnotation(Marker::class.java)
            val a2 = method.getAnnotation(Marker::class.java)
            __check((a1?.kind).toString(), "value")
            __check((a2?.kind).toString(), "action")
        }
