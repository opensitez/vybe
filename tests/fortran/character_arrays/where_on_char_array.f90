! vybe-test: fortran/character_arrays/where_on_char_array
! origin: languages/fortran/tests/fortran/test_character_arrays.rs

program test
    character(len=5) :: src(4) = ['apple', 'mango', 'grape', 'pear ']
    character(len=5) :: dst(4)
    dst = '     '
    where (src /= 'mango')
        dst = src
    end where
    print *, trim(dst(1))
    print *, trim(dst(2))
end program test
