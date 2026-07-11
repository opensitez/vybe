
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_flock_shared, "require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_SH)", "0");
ruby_test!(test_flock_exclusive, "require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_EX)", "0");
ruby_test!(test_flock_unlock, "require 'tempfile'; t = Tempfile.new('lock'); t.flock(File::LOCK_EX); puts t.flock(File::LOCK_UN)", "0");
ruby_test!(test_flock_nonblock, "require 'tempfile'; t = Tempfile.new('lock'); puts t.flock(File::LOCK_EX | File::LOCK_NB)", "0");
ruby_test!(test_flock_conflict, "require 'tempfile'; t1 = Tempfile.new('lock'); t2 = File.open(t1.path, 'r'); t1.flock(File::LOCK_EX); puts t2.flock(File::LOCK_EX | File::LOCK_NB)", "false"); // returns false if locked by another descriptor
ruby_test!(test_flock_error_closed, "require 'tempfile'; t = Tempfile.new('lock'); t.close; begin; t.flock(File::LOCK_EX); rescue IOError; puts 'err'; end", "err");
