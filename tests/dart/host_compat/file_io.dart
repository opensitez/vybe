// vybe-test: dart/host_compat/file_io
// origin: languages/dart/tests/dart/test_host_compat.rs

var content = File.readAsStringSync('data.txt');
File.writeAsStringSync('output.txt', content);

void main() {}
