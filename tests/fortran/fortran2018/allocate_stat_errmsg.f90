! vybe-test: fortran/fortran2018/allocate_stat_errmsg
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer, allocatable :: a(:)
    integer :: stat
    character(len=100) :: errmsg
    allocate(a(10), stat=stat, errmsg=errmsg)
    if (stat /= 0) then
        print *, trim(errmsg)
    else
        print *, size(a)
    end if
    deallocate(a)
end program test
