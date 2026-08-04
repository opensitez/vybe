! vybe-test: fortran/fortran_coverage_bulk/fortran_bulk_nested_associate_array_aliasing
! origin: languages/fortran/tests/fortran/test_fortran_coverage_bulk.rs

program fortran_bulk_nested_associate_array_aliasing
    integer :: buffer(4)
    integer :: base

    base = 5

    associate (whole => buffer)
        whole = (/ 1, 2, 3, 4 /)
        associate (edge => whole(2:3))
            edge = edge + base
        end associate
        print *, whole(1), whole(2), whole(3), whole(4)
    end associate
end program fortran_bulk_nested_associate_array_aliasing
