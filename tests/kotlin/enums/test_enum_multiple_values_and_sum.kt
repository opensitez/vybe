// vybe-test: kotlin/enums/test_enum_multiple_values_and_sum
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Piece { A, B, C, D }
fun main() { var total = 0
for (p in arrayOf(Piece.A, Piece.B, Piece.C, Piece.D)) { total += p }
println(total) }

