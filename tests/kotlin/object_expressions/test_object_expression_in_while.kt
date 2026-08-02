// vybe-test: kotlin/object_expressions/test_object_expression_in_while
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun main() { val o = object { var n = 0
fun next() { n += 1 } }
var i = 0
while (i < 2) { o.next()
i += 1 }
println(o.n) }

