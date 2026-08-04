// vybe-test: dart/control_flow_advanced/switch_multiple_values
// origin: languages/dart/tests/dart/test_control_flow_advanced.rs

void main() {
  var x = 'b';
  switch (x) {
    case 'a':
    case 'b':
      print('vowel-ish');
      break;
    default:
      print('other');
  }
}