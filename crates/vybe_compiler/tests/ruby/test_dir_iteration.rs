
macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    }
}

ruby_test!(test_dir_iteration_each, "Dir.mkdir('/tmp/test_dir_iter'); File.write('/tmp/test_dir_iter/a', 'a'); acc = []; Dir.new('/tmp/test_dir_iter').each { |f| acc << f }; puts acc.sort.join('-'); File.delete('/tmp/test_dir_iter/a'); Dir.rmdir('/tmp/test_dir_iter')", ".-..-a");
ruby_test!(test_dir_iteration_each_child, "Dir.mkdir('/tmp/test_dir_iter2'); File.write('/tmp/test_dir_iter2/a', 'a'); acc = []; Dir.new('/tmp/test_dir_iter2').each_child { |f| acc << f }; puts acc.join('-'); File.delete('/tmp/test_dir_iter2/a'); Dir.rmdir('/tmp/test_dir_iter2')", "a");
ruby_test!(test_dir_iteration_children, "Dir.mkdir('/tmp/test_dir_iter3'); File.write('/tmp/test_dir_iter3/a', 'a'); puts Dir.new('/tmp/test_dir_iter3').children.join('-'); File.delete('/tmp/test_dir_iter3/a'); Dir.rmdir('/tmp/test_dir_iter3')", "a");
ruby_test!(test_dir_class_each_child, "Dir.mkdir('/tmp/test_dir_iter4'); File.write('/tmp/test_dir_iter4/a', 'a'); acc = []; Dir.each_child('/tmp/test_dir_iter4') { |f| acc << f }; puts acc.join('-'); File.delete('/tmp/test_dir_iter4/a'); Dir.rmdir('/tmp/test_dir_iter4')", "a");
ruby_test!(test_dir_class_children, "Dir.mkdir('/tmp/test_dir_iter5'); File.write('/tmp/test_dir_iter5/a', 'a'); puts Dir.children('/tmp/test_dir_iter5').join('-'); File.delete('/tmp/test_dir_iter5/a'); Dir.rmdir('/tmp/test_dir_iter5')", "a");
