! vybe-test: fortran/fortran_coverage_bulk/fortran_bulk_name_resolution_nested_blocks
! origin: languages/fortran/tests/fortran/test_fortran_coverage_bulk.rs

program fortran_bulk_name_resolution_nested_blocks
    integer :: outer
    outer = 10
    block
        integer :: outer
        outer = 20
        block
            integer :: inner
            inner = outer + 1
            if ((inner) /= 21) then
    print *, "FAIL: want [21] got [", inner, "]"
    stop 1
end if
        end block
        if ((outer) /= 20) then
    print *, "FAIL: want [20] got [", outer, "]"
    stop 1
end if
    end block
    if ((outer) /= 10) then
    print *, "FAIL: want [10] got [", outer, "]"
    stop 1
end if
end program fortran_bulk_name_resolution_nested_blocks
