// vybe-test: kotlin/kotlin_sealed_class_hierarchy/test_sealed_types_prevent_external_extension
// origin: languages/kotlin/tests/kotlin/test_kotlin_sealed_class_hierarchy.rs

sealed class Response {
            object Ok : Response()
            data class Error(val code: Int) : Response()
        }

        fun message(response: Response): String = when (response) {
            is Response.Ok -> "ok"
            is Response.Error -> "err=" + response.code
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((message(Response.Ok)).toString(), "ok")
            __check((message(Response.Error(5))).toString(), "err=5")
        }
