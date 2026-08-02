// vybe-test: kotlin/comments/test_comment_inside_raw_string_body
// origin: languages/kotlin/tests/kotlin/test_comments.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """line1
// not parsed as comment
line3"""
            __check((text.lines().size).toString(), "3")
        }
