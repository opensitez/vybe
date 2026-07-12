macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    test_fcntl_basic,
    "require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); puts t.fcntl(Fcntl::F_GETFD).is_a?(Integer)",
    "true"
);
ruby_test!(
    test_fcntl_setfd,
    "require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); flags = t.fcntl(Fcntl::F_GETFD); puts t.fcntl(Fcntl::F_SETFD, flags | Fcntl::FD_CLOEXEC)",
    "0"
); // usually returns 0 on success
ruby_test!(
    test_fcntl_error,
    "require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); t.close; begin; t.fcntl(Fcntl::F_GETFD); rescue IOError; puts 'err'; end",
    "err"
);
ruby_test!(
    test_fcntl_invalid_cmd,
    "require 'tempfile'; require 'fcntl'; t = Tempfile.new('fcntl'); begin; t.fcntl(99999); rescue Errno::EINVAL; puts 'err'; end",
    "err"
);
