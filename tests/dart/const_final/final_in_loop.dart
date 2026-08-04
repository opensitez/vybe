// vybe-test: dart/const_final/final_in_loop
// origin: languages/dart/tests/dart/test_const_final.rs

void main() { for (var i = 0; i < 3; i++) { final v = i * 2; print(v); } }