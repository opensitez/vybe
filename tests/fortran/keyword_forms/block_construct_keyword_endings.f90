! vybe-test: fortran/keyword_forms/block_construct_keyword_endings
! origin: languages/fortran/tests/fortran/test_keyword_forms.rs

program kw_blocks
    implicit none
    integer :: a(1:3)
    integer :: i

    do i = 1, 3
        a(i) = i
    end do

    do i = 1, 3
        if (a(i) == 2) then
            a(i) = a(i) + 1
        else if (a(i) == 3) then
            a(i) = a(i) - 1
        else
            a(i) = a(i)
        end if
end do

    select case (a(2))
    case (1:2)
        a(2) = a(2) * 2
    case default
        a(2) = 0
    end select

    where (a > 0)
        a = a + 1
    end where
end program kw_blocks
