use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(dimensional_rectangular_length, "int[][] m = new int[2][3]; System.out.println(m.length);", "2");
jt!(dimensional_row_length, "int[][] m = new int[2][3]; System.out.println(m[0].length);", "3");
jt!(dimensional_literal_shape, "int[][] m = {{1,2},{3,4,5}}; System.out.println(m[0].length);", "2");
jt!(dimensional_literal_second, "int[][] m = {{1,2},{3,4,5}}; System.out.println(m[1].length);", "3");
jt!(dimensional_sum_elements, "int[][] m = {{1,2},{3,4}}; int s = 0; for(int i = 0; i < m.length; i++) { for(int j = 0; j < m[i].length; j++) s += m[i][j]; } System.out.println(s);", "10");
jt!(dimensional_access_bottom_right, "int[][] m = {{1,2,3},{4,5,6}}; System.out.println(m[1][2]);", "6");
jt!(dimensional_assign_then_read, "int[][] m = new int[2][2]; m[1][0] = 9; System.out.println(m[1][0]);", "9");
jt!(dimensional_nested_sum_by_rows, "int[][] m = {{1,1,1},{2,2,2}}; int total = 0; for(int i = 0; i < m.length; i++) { int rowSum = 0; for(int j = 0; j < m[i].length; j++) rowSum += m[i][j]; total += rowSum; } System.out.println(total);", "9");
jt!(dimensional_jagged_access, "int[][] m = new int[2][]; m[0] = new int[]{1}; m[1] = new int[]{2,3}; System.out.println(m[1][0] + m[1][1]);", "5");
jt!(dimensional_initialize_and_modify, "int[][] m = {{1,2,3},{4,5,6}}; m[0][1] = 9; System.out.println(m[0][1]);", "9");
jt!(dimensional_default_row_values, "int[][] m = new int[2][3]; System.out.println(m[0][0] + m[1][2]);", "0");
jt!(dimensional_boolean_matrix, "boolean[][] b = {{true, false}, {false, true}}; int c = 0; if(b[0][0]) c++; if(b[1][1]) c++; System.out.println(c);", "2");
jt!(three_dimensional_array_shape, "int[][][] t = new int[2][1][3]; System.out.println(t.length + t[0].length + t[0][0].length);", "6");
jt!(three_dimensional_access, "int[][][] t = {{{1,2},{3,4}},{{5,6}}}; System.out.println(t[1][0][1]);", "6");
jt!(dimensional_string_matrix, "String[][] names = {{\"a\",\"b\"},{\"c\",\"d\"}}; System.out.println(names[0][1] + names[1][0]);", "bc");
jt!(dimensional_sum_filtered_row, "int[][] m = {{1,2,3},{4,5,6}}; int s = 0; for(int j = 0; j < m[0].length; j++) s += m[0][j]; System.out.println(s);", "6");
jt!(dimensional_count_rows_with_len, "int[][] m = {{1},{2,3},{4,5,6}}; int c = 0; for(int i = 0; i < m.length; i++) c += m[i].length; System.out.println(c);", "6");
jt!(dimensional_mutate_in_nested_for, "int[][] m = {{1,1,1},{1,1,1}}; for(int i = 0; i < m.length; i++) for(int j = 0; j < m[i].length; j++) m[i][j] = i + j; System.out.println(m[1][2]);", "3");
jt!(dimensional_max_element_first_row, "int[][] m = {{8,1,5},{2,7}}; int max = m[0][0]; for(int j = 0; j < m[0].length; j++) if(m[0][j] > max) max = m[0][j]; System.out.println(max);", "8");
jt!(dimensional_min_element_last_row, "int[][] m = {{8,1,5},{2,7}}; int min = m[1][0]; for(int j = 0; j < m[1].length; j++) if(m[1][j] < min) min = m[1][j]; System.out.println(min);", "2");
jt!(dimensional_index_calculation, "int[][] m = {{1,2,3},{4,5,6}}; int i = 1; int j = 1; System.out.println(m[i][j]);", "5");
jt!(dimensional_mixed_refs, "Object[] a = {new int[]{1,2}, new int[]{3,4}}; System.out.println(a.length);", "2");
jt!(dimensional_flat_access_sum, "int[][] m = {{1,2},{3,4},{5,6}}; int s = m[0][0] + m[1][1] + m[2][1]; System.out.println(s);", "11");
jt!(dimensional_string_length_in_cells, "String[][] words = {{\"hi\", \"ok\"}, {\"x\"}}; System.out.println(words[1][0].length());", "1");
jt!(dimensional_compare_row_refs, "int[][] m = {{1,2,3}}; int[][] n = m; System.out.println(n[0][2]);", "3");
