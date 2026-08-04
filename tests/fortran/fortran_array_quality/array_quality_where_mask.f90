! vybe-test: fortran/fortran_array_quality/array_quality_where_mask
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_where_mask
    integer, dimension(6) :: source
    integer, dimension(6) :: target
    source = (/ 1, 2, 3, 4, 5, 6 /)
    where (mod(source,2) == 0)
        target = source * 2
    elsewhere
        target = source
    end where
    print *, target(2), target(3)
end program array_quality_where_mask
