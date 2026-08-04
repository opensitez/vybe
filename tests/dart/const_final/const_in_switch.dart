// vybe-test: dart/const_final/const_in_switch
// origin: languages/dart/tests/dart/test_const_final.rs

const kMax = 3;
void main() {
  switch (kMax) {
    case 3: print('three'); break;
    default: print('other');
  }
}