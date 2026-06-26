//! Dart `Uri`: parse, components, query/fragment, replace, normalize, origin.

dart_cases! {
    uri_parse_http_scheme => {
        r#"void main() {
  var u = Uri.parse('http://example.com/');
  print(u.scheme);
}"#,
        ["http"]
    };

    uri_parse_https_scheme => {
        r#"void main() {
  var u = Uri.parse('https://secure.example.org/');
  print(u.scheme);
}"#,
        ["https"]
    };

    uri_parse_host_from_http_url => {
        r#"void main() {
  var u = Uri.parse('http://api.example.com/v1');
  print(u.host);
}"#,
        ["api.example.com"]
    };

    uri_parse_host_localhost => {
        r#"void main() {
  var u = Uri.parse('http://localhost:3000/');
  print(u.host);
}"#,
        ["localhost"]
    };

    uri_parse_explicit_port => {
        r#"void main() {
  var u = Uri.parse('http://example.com:8080/path');
  print(u.port);
}"#,
        ["8080"]
    };

    uri_parse_default_http_port => {
        r#"void main() {
  var u = Uri.parse('http://example.com/path');
  print(u.port);
}"#,
        ["80"]
    };

    uri_parse_default_https_port => {
        r#"void main() {
  var u = Uri.parse('https://example.com/path');
  print(u.port);
}"#,
        ["443"]
    };

    uri_parse_root_path => {
        r#"void main() {
  var u = Uri.parse('http://example.com/');
  print(u.path);
}"#,
        ["/"]
    };

    uri_parse_multi_segment_path => {
        r#"void main() {
  var u = Uri.parse('http://example.com/api/users/42');
  print(u.path);
}"#,
        ["/api/users/42"]
    };

    uri_parse_path_without_trailing_slash => {
        r#"void main() {
  var u = Uri.parse('http://example.com/data.json');
  print(u.path);
}"#,
        ["/data.json"]
    };

    uri_parse_single_query_parameter => {
        r#"void main() {
  var u = Uri.parse('http://example.com/search?q=dart');
  print(u.queryParameters['q']);
}"#,
        ["dart"]
    };

    uri_parse_multiple_query_parameters => {
        r#"void main() {
  var u = Uri.parse('http://example.com/?a=1&b=two');
  print(u.queryParameters['a']);
  print(u.queryParameters['b']);
}"#,
        ["1", "two"]
    };

    uri_parse_query_parameter_with_empty_value => {
        r#"void main() {
  var u = Uri.parse('http://example.com/?flag=');
  print(u.queryParameters['flag']);
  print(u.queryParameters.containsKey('flag'));
}"#,
        ["", "true"]
    };

    uri_parse_percent_encoded_query_value => {
        r#"void main() {
  var u = Uri.parse('http://example.com/?msg=hello%20world');
  print(u.queryParameters['msg']);
}"#,
        ["hello world"]
    };

    uri_parse_fragment => {
        r#"void main() {
  var u = Uri.parse('http://example.com/docs#intro');
  print(u.fragment);
}"#,
        ["intro"]
    };

    uri_parse_no_fragment_is_empty => {
        r#"void main() {
  var u = Uri.parse('http://example.com/page');
  print(u.fragment);
  print(u.hasFragment);
}"#,
        ["", "false"]
    };

    uri_has_scheme_true_for_absolute_http => {
        r#"void main() {
  var u = Uri.parse('http://example.com/');
  print(u.hasScheme);
}"#,
        ["true"]
    };

    uri_has_scheme_false_for_relative_path => {
        r#"void main() {
  var u = Uri.parse('/relative/path');
  print(u.hasScheme);
}"#,
        ["false"]
    };

    uri_has_authority_true_for_http => {
        r#"void main() {
  var u = Uri.parse('http://example.com/path');
  print(u.hasAuthority);
}"#,
        ["true"]
    };

    uri_has_authority_false_for_relative => {
        r#"void main() {
  var u = Uri.parse('relative/file.txt');
  print(u.hasAuthority);
}"#,
        ["false"]
    };

    uri_origin_http_without_explicit_port => {
        r#"void main() {
  var u = Uri.parse('http://example.com/path');
  print(u.origin);
}"#,
        ["http://example.com"]
    };

    uri_origin_https_default_port_omitted => {
        r#"void main() {
  var u = Uri.parse('https://example.com:443/secure');
  print(u.origin);
}"#,
        ["https://example.com"]
    };

    uri_origin_includes_non_default_port => {
        r#"void main() {
  var u = Uri.parse('https://example.com:8443/app');
  print(u.origin);
}"#,
        ["https://example.com:8443"]
    };

    uri_replace_scheme => {
        r#"void main() {
  var u = Uri.parse('http://example.com/a').replace(scheme: 'https');
  print(u.scheme);
  print(u.host);
}"#,
        ["https", "example.com"]
    };

    uri_replace_host => {
        r#"void main() {
  var u = Uri.parse('http://old.com/x').replace(host: 'new.com');
  print(u.host);
  print(u.path);
}"#,
        ["new.com", "/x"]
    };

    uri_replace_port => {
        r#"void main() {
  var u = Uri.parse('http://example.com/').replace(port: 9090);
  print(u.port);
}"#,
        ["9090"]
    };

    uri_replace_path => {
        r#"void main() {
  var u = Uri.parse('http://example.com/old').replace(path: '/new');
  print(u.path);
}"#,
        ["/new"]
    };

    uri_replace_query => {
        r#"void main() {
  var u = Uri.parse('http://example.com/').replace(query: 'x=1&y=2');
  print(u.query);
  print(u.queryParameters['y']);
}"#,
        ["x=1&y=2", "2"]
    };

    uri_replace_fragment => {
        r#"void main() {
  var u = Uri.parse('http://example.com/doc').replace(fragment: 'section');
  print(u.fragment);
  print(u.hasFragment);
}"#,
        ["section", "true"]
    };

    uri_normalize_path_collapses_parent_segment => {
        r#"void main() {
  var u = Uri.parse('http://example.com/a/b/../c').normalizePath();
  print(u.path);
}"#,
        ["/a/c"]
    };

    uri_normalize_path_removes_current_directory_segment => {
        r#"void main() {
  var u = Uri.parse('http://example.com/a/./b').normalizePath();
  print(u.path);
}"#,
        ["/a/b"]
    };

    uri_normalize_path_on_simple_path_unchanged => {
        r#"void main() {
  var u = Uri.parse('http://example.com/foo/bar').normalizePath();
  print(u.path);
}"#,
        ["/foo/bar"]
    };

    uri_to_string_roundtrip_preserves_http_url => {
        r#"void main() {
  var original = 'http://example.com/api?k=v#top';
  var u = Uri.parse(original);
  print(u.toString());
}"#,
        ["http://example.com/api?k=v#top"]
    };

    uri_to_string_roundtrip_https_with_port => {
        r#"void main() {
  var original = 'https://example.com:8443/data';
  var u = Uri.parse(original);
  print(u.toString());
}"#,
        ["https://example.com:8443/data"]
    };

    uri_authority_includes_host_and_port => {
        r#"void main() {
  var u = Uri.parse('http://example.com:8080/path');
  print(u.authority);
}"#,
        ["example.com:8080"]
    };

    uri_is_absolute_true_for_http => {
        r#"void main() {
  var u = Uri.parse('http://example.com/');
  print(u.isAbsolute);
}"#,
        ["true"]
    };

    uri_is_absolute_false_for_relative_path => {
        r#"void main() {
  var u = Uri.parse('/only/path');
  print(u.isAbsolute);
}"#,
        ["false"]
    };

    uri_path_segments_split_on_slashes => {
        r#"void main() {
  var u = Uri.parse('http://example.com/a/b/c');
  print(u.pathSegments.join(','));
}"#,
        ["a,b,c"]
    };

    uri_has_query_true_when_query_present => {
        r#"void main() {
  var u = Uri.parse('http://example.com/?q=1');
  print(u.hasQuery);
}"#,
        ["true"]
    };

    uri_has_query_false_without_query => {
        r#"void main() {
  var u = Uri.parse('http://example.com/path');
  print(u.hasQuery);
}"#,
        ["false"]
    };

    uri_has_empty_path_true_for_authority_only => {
        r#"void main() {
  var u = Uri.parse('http://example.com');
  print(u.hasEmptyPath);
}"#,
        ["true"]
    };

    uri_user_info_in_authority => {
        r#"void main() {
  var u = Uri.parse('http://user:pass@example.com/');
  print(u.userInfo);
}"#,
        ["user:pass"]
    };

    uri_http_constructor_builds_expected_path => {
        r#"void main() {
  var u = Uri.http('example.com', '/api');
  print(u.scheme);
  print(u.host);
  print(u.path);
}"#,
        ["http", "example.com", "/api"]
    };

    uri_https_constructor_with_query_parameters => {
        r#"void main() {
  var u = Uri.https('example.com', '/search', {'q': 'dart'});
  print(u.scheme);
  print(u.queryParameters['q']);
}"#,
        ["https", "dart"]
    };

    uri_https_constructor_with_explicit_port => {
        r#"void main() {
  var u = Uri.https('example.com', '/', null, 8443);
  print(u.port);
  print(u.scheme);
}"#,
        ["8443", "https"]
    };

    uri_file_constructor_scheme => {
        r#"void main() {
  var u = Uri.file('/tmp/data.txt');
  print(u.scheme);
  print(u.path.endsWith('data.txt'));
}"#,
        ["file", "true"]
    };

    uri_replace_clears_fragment_with_empty_string => {
        r#"void main() {
  var u = Uri.parse('http://example.com/#old').replace(fragment: '');
  print(u.fragment);
  print(u.hasFragment);
}"#,
        ["", "false"]
    };

    uri_replace_clears_query_with_empty_string => {
        r#"void main() {
  var u = Uri.parse('http://example.com/?a=1').replace(query: '');
  print(u.query);
  print(u.hasQuery);
}"#,
        ["", "false"]
    };

    uri_replace_path_segments => {
        r#"void main() {
  var u = Uri.parse('http://example.com/old/x').replace(pathSegments: ['new', 'y']);
  print(u.path);
}"#,
        ["/new/y"]
    };

    uri_resolve_relative_path_against_base => {
        r#"void main() {
  var base = Uri.parse('http://example.com/a/b/');
  var resolved = base.resolve('c');
  print(resolved.path);
}"#,
        ["/a/b/c"]
    };

    uri_resolve_parent_relative_segment => {
        r#"void main() {
  var base = Uri.parse('http://example.com/a/b/c');
  var resolved = base.resolve('../d');
  print(resolved.path);
}"#,
        ["/a/b/d"]
    };

    uri_resolve_uri_object => {
        r#"void main() {
  var base = Uri.parse('http://example.com/base/');
  var rel = Uri.parse('item');
  var resolved = base.resolveUri(rel);
  print(resolved.path);
}"#,
        ["/base/item"]
    };

    uri_query_empty_when_no_parameters => {
        r#"void main() {
  var u = Uri.parse('http://example.com/path');
  print(u.query);
  print(u.queryParameters.isEmpty);
}"#,
        ["", "true"]
    };

    uri_percent_encoded_path_segment_decoded_in_path => {
        r#"void main() {
  var u = Uri.parse('http://example.com/a%20b');
  print(u.path);
}"#,
        ["/a b"]
    };

    uri_host_is_empty_for_relative_uri => {
        r#"void main() {
  var u = Uri.parse('just/a/path');
  print(u.host);
  print(u.hasAuthority);
}"#,
        ["", "false"]
    };
}
