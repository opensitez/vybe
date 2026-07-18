use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(sum_2x2, "int[][] m = {{1,2},{3,4}}; int s = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) s += m[i][j]; } System.out.println(s);", "10");
jt!(identity_2x2_sum, "int[][] m = {{1,0},{0,1}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) s += m[i][j]; System.out.println(s);", "2");
jt!(diag_sum, "int[][] m = {{1,2,3},{4,5,6},{7,8,9}}; int s = 0; for (int i = 0; i < m.length; i++) s += m[i][i]; System.out.println(s);", "15");
jt!(upper_tri_sum, "int[][] m = {{1,2,3},{4,5,6},{7,8,9}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = i; j < m[i].length; j++) s += m[i][j]; System.out.println(s);", "30");
jt!(lower_tri_sum, "int[][] m = {{1,2,3},{4,5,6},{7,8,9}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = 0; j <= i; j++) s += m[i][j]; System.out.println(s);", "35");
jt!(row_sums, "int[][] m = {{1,2},{3,4},{5,6}}; int total = 0; for (int i = 0; i < m.length; i++) { int row = 0; for (int j = 0; j < m[i].length; j++) row += m[i][j]; total += row; } System.out.println(total);", "21");
jt!(transpose_value, "int[][] m = {{1,2,3},{4,5,6}}; int v = m[1][0] + m[0][2]; System.out.println(v);", "7");
jt!(search_in_matrix, "int[][] m = {{1,2},{3,4}}; int r = -1; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) { if (m[i][j] == 3) r = i * 10 + j; } } System.out.println(r);", "10");
jt!(mutate_matrix, "int[][] m = {{1,2},{3,4}}; for (int i = 0; i < m.length; i++) m[i][i] = 0; System.out.println(m[0][0] + m[1][1]);", "0");
jt!(multiply_matrix_scalar, "int[][] m = {{1,2},{3,4}}; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) m[i][j] *= 2; System.out.println(m[1][1]);", "8");
jt!(add_matrix_rows, "int[][] m = {{1,1},{2,2}}; int a = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) if (i == 0) a += m[i][j]; } System.out.println(a);", "2");
jt!(column_sum, "int[][] m = {{1,2,3},{4,5,6},{7,8,9}}; int c = 0; for (int i = 0; i < m.length; i++) c += m[i][1]; System.out.println(c);", "15");
jt!(jagged_sum, "int[][] m = {{1,2,3},{4,5},{6}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) s += m[i][j]; System.out.println(s);", "21");
jt!(jagged_first_last, "int[][] m = {{1},{2,3},{4,5,6}}; int s = m[0][0] + m[2][2]; System.out.println(s);", "7");
jt!(sum_if_even_pos, "int[][] m = {{1,2,3},{4,5,6}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) if (((i + j) & 1) == 0) s += m[i][j]; System.out.println(s);", "11");
jt!(count_cells, "int[][] m = {{1,2,3},{4},{5,6,7}}; int c = 0; for (int i = 0; i < m.length; i++) c += m[i].length; System.out.println(c);", "6");
jt!(mirror_first, "int[][] m = {{1,2,3},{4,5,6}}; int s = m[0][0] + m[0][2] + m[1][0] + m[1][2]; System.out.println(s);", "12");
jt!(walk_while_cells, "int[][] m = {{1,2},{3,4,5}}; int s = 0; int i = 0; while (i < m.length) { int j = 0; while (j < m[i].length) { s += m[i][j]; j++; } i++; } System.out.println(s);", "15");
jt!(for_each_rows, "int[][] m = {{1,2},{3,4}}; int s = 0; for (int[] row : m) for (int v : row) s += v; System.out.println(s);", "10");
jt!(compare_row0_row1, "int[][] m = {{1,2},{3,4}}; boolean ok = m[0][1] < m[1][1]; System.out.println(ok);", "true");
jt!(flip_sign_odd_rows, "int[][] m = {{1,-2},{3,-4}}; for (int i = 0; i < m.length; i++) if ((i & 1) == 1) for (int j = 0; j < m[i].length; j++) m[i][j] = -m[i][j]; System.out.println(m[1][1]);", "4");
jt!(sum_top_row, "int[][] m = {{9,8,7},{6,5,4}}; int s = 0; for (int i = 0; i < m[0].length; i++) s += m[0][i]; System.out.println(s);", "24");
jt!(sum_bottom_row, "int[][] m = {{9,8,7},{6,5,4}}; int s = 0; for (int i = 0; i < m[1].length; i++) s += m[1][i]; System.out.println(s);", "15");
jt!(first_col_sum, "int[][] m = {{1,2,3},{4,5,6}}; int s = 0; for (int i = 0; i < m.length; i++) s += m[i][0]; System.out.println(s);", "5");
jt!(last_col_sum, "int[][] m = {{1,2,3},{4,5,6}}; int s = 0; for (int i = 0; i < m.length; i++) s += m[i][m[i].length - 1]; System.out.println(s);", "9");
jt!(flatten_sum, "int[][] m = {{1,2},{3,4},{5}}; int s = 0; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) s += m[i][j]; System.out.println(s);", "15");
jt!(transpose_virtual, "int[][] m = {{1,2,3},{4,5,6}}; int s = m[0][1] + m[1][0]; System.out.println(s);", "6");
jt!(matrix_boolean_or, "boolean[][] m = {{true,false},{false,true}}; boolean ok = false; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) ok = ok || m[i][j]; System.out.println(ok);", "true");
jt!(matrix_min, "int[][] m = {{9,8},{6,7}}; int min = m[0][0]; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) if (m[i][j] < min) min = m[i][j]; System.out.println(min);", "6");
jt!(matrix_max, "int[][] m = {{9,8},{6,7}}; int max = m[0][0]; for (int i = 0; i < m.length; i++) for (int j = 0; j < m[i].length; j++) if (m[i][j] > max) max = m[i][j]; System.out.println(max);", "9");

