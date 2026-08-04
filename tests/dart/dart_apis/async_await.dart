// vybe-test: dart/dart_apis/async_await
// origin: languages/dart/tests/dart/test_dart_apis.rs

class App { fetchData() async { return 'data'; } main() async { var d = await fetchData(); print(d); } }