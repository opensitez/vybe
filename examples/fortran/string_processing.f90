! String Processing — text parsing, tokenization, pattern matching
! Covers: character variables, len_trim, adjustl/adjustr, index, scan, verify,
!         character arrays, internal files (write/read to string), 
!         variable-length strings via allocatable character.

module string_utils
    implicit none
    private

    public :: to_upper, to_lower, trim_all, split, join, &
              starts_with, ends_with, contains_str, replace_str, &
              count_occurrences, reverse_str, is_integer, is_real, &
              parse_csv_line, word_count

contains

    elemental pure function to_upper(c) result(u)
        character(len=1), intent(in) :: c
        character(len=1) :: u
        integer :: code
        code = iachar(c)
        if (code >= iachar('a') .and. code <= iachar('z')) then
            u = achar(code - 32)
        else
            u = c
        end if
    end function to_upper

    elemental pure function to_lower(c) result(l)
        character(len=1), intent(in) :: c
        character(len=1) :: l
        integer :: code
        code = iachar(c)
        if (code >= iachar('A') .and. code <= iachar('Z')) then
            l = achar(code + 32)
        else
            l = c
        end if
    end function to_lower

    ! Convert entire string to upper/lower case
    pure function str_upper(s) result(u)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: u
        integer :: i
        do i = 1, len(s)
            u(i:i) = to_upper(s(i:i))
        end do
    end function str_upper

    pure function str_lower(s) result(l)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: l
        integer :: i
        do i = 1, len(s)
            l(i:i) = to_lower(s(i:i))
        end do
    end function str_lower

    ! Remove all whitespace (leading, trailing, internal)
    pure function trim_all(s) result(t)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: t
        integer :: i, j
        j = 0
        do i = 1, len_trim(s)
            if (s(i:i) /= ' ' .and. s(i:i) /= char(9)) then
                j = j + 1
                t(j:j) = s(i:i)
            end if
        end do
        t(j+1:) = ' '
    end function trim_all

    logical pure function starts_with(s, prefix)
        character(len=*), intent(in) :: s, prefix
        integer :: lp
        lp = len_trim(prefix)
        starts_with = (len_trim(s) >= lp) .and. (s(1:lp) == prefix(1:lp))
    end function starts_with

    logical pure function ends_with(s, suffix)
        character(len=*), intent(in) :: s, suffix
        integer :: ls, lsuf
        ls   = len_trim(s)
        lsuf = len_trim(suffix)
        ends_with = (ls >= lsuf) .and. (s(ls-lsuf+1:ls) == suffix(1:lsuf))
    end function ends_with

    logical pure function contains_str(s, sub)
        character(len=*), intent(in) :: s, sub
        contains_str = (index(s, trim(sub)) > 0)
    end function contains_str

    pure function count_occurrences(s, sub) result(n)
        character(len=*), intent(in) :: s, sub
        integer :: n, pos, start, lsub
        n    = 0
        lsub = len_trim(sub)
        if (lsub == 0) return
        start = 1
        do
            pos = index(s(start:), trim(sub))
            if (pos == 0) exit
            n     = n + 1
            start = start + pos + lsub - 1
        end do
    end function count_occurrences

    pure function replace_str(s, old, new) result(r)
        character(len=*), intent(in) :: s, old, new
        character(len=len(s)*3) :: r   ! generous upper bound
        integer :: pos, start, lold, lnew, rlen
        lold  = len_trim(old)
        lnew  = len_trim(new)
        r     = ' '
        rlen  = 0
        start = 1
        do
            pos = index(s(start:), trim(old))
            if (pos == 0) then
                ! Copy remainder
                r(rlen+1:rlen+len_trim(s)-start+1) = s(start:len_trim(s))
                exit
            end if
            ! Copy up to match
            r(rlen+1:rlen+pos-1) = s(start:start+pos-2)
            rlen = rlen + pos - 1
            ! Insert replacement
            r(rlen+1:rlen+lnew) = new(1:lnew)
            rlen  = rlen + lnew
            start = start + pos + lold - 1
        end do
    end function replace_str

    pure function reverse_str(s) result(r)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: r
        integer :: i, n
        n = len_trim(s)
        do i = 1, n
            r(i:i) = s(n-i+1:n-i+1)
        end do
        r(n+1:) = ' '
    end function reverse_str

    logical pure function is_integer(s)
        character(len=*), intent(in) :: s
        character(len=len_trim(s)) :: t
        integer :: i, start
        t = adjustl(s)
        is_integer = .false.
        if (len_trim(t) == 0) return
        start = 1
        if (t(1:1) == '+' .or. t(1:1) == '-') start = 2
        if (start > len_trim(t)) return
        do i = start, len_trim(t)
            if (scan(t(i:i), '0123456789') == 0) return
        end do
        is_integer = .true.
    end function is_integer

    logical pure function is_real(s)
        character(len=*), intent(in) :: s
        character(len=len_trim(s)) :: t
        integer :: i, start, dot_count, mantissa_digits, exponent_digits
        logical :: seen_exponent
        t = adjustl(s)
        is_real = .false.
        if (len_trim(t) == 0) return
        start = 1
        dot_count = 0
        mantissa_digits = 0
        exponent_digits = 0
        seen_exponent = .false.
        if (t(1:1) == '+' .or. t(1:1) == '-') start = 2
        if (start > len_trim(t)) return
        do i = start, len_trim(t)
            if (scan(t(i:i), '0123456789') > 0) then
                if (seen_exponent) then
                    exponent_digits = exponent_digits + 1
                else
                    mantissa_digits = mantissa_digits + 1
                end if
            else if (t(i:i) == '.') then
                if (seen_exponent) return
                dot_count = dot_count + 1
                if (dot_count > 1) return
            else if (scan(t(i:i), 'eEdD') > 0) then
                if (seen_exponent .or. mantissa_digits == 0 .or. i == len_trim(t)) return
                seen_exponent = .true.
            else if (t(i:i) == '+' .or. t(i:i) == '-') then
                if (.not. seen_exponent .or. i == len_trim(t)) return
                if (scan(t(i-1:i-1), 'eEdD') == 0) return
            else
                return
            end if
        end do
        if (mantissa_digits == 0) return
        if (seen_exponent .and. exponent_digits == 0) return
        is_real = (dot_count > 0 .or. seen_exponent)
    end function is_real

    ! Split string by delimiter into tokens array
    subroutine split(s, delim, tokens, n_tokens)
        character(len=*), intent(in)  :: s, delim
        character(len=256), allocatable, intent(out) :: tokens(:)
        integer, intent(out) :: n_tokens

        integer :: pos, start, count, ldelim
        character(len=len(s)) :: work

        ldelim = len_trim(delim)
        work   = s
        count  = 1
        start  = 1

        ! Count tokens
        do
            pos = index(work(start:), trim(delim))
            if (pos == 0) exit
            count = count + 1
            start = start + pos + ldelim - 1
        end do

        allocate(tokens(count))
        n_tokens = count
        start = 1
        count = 0

        do
            pos = index(work(start:), trim(delim))
            count = count + 1
            if (pos == 0) then
                tokens(count) = adjustl(work(start:len_trim(work)))
                exit
            end if
            tokens(count) = adjustl(work(start:start+pos-2))
            start = start + pos + ldelim - 1
        end do
    end subroutine split

    ! Join array of strings with separator
    function join(tokens, n, sep) result(s)
        integer, intent(in) :: n
        character(len=*), intent(in) :: tokens(n), sep
        character(len=4096) :: s
        integer :: i, pos
        s   = ' '
        pos = 1
        do i = 1, n
            s(pos:pos+len_trim(tokens(i))-1) = tokens(i)(1:len_trim(tokens(i)))
            pos = pos + len_trim(tokens(i))
            if (i < n) then
                s(pos:pos+len_trim(sep)-1) = sep(1:len_trim(sep))
                pos = pos + len_trim(sep)
            end if
        end do
    end function join

    ! Parse a CSV line into fields
    subroutine parse_csv_line(line, fields, n_fields)
        character(len=*), intent(in) :: line
        character(len=256), allocatable, intent(out) :: fields(:)
        integer, intent(out) :: n_fields
        logical :: in_quotes
        integer :: i, start, count
        character(len=len(line)) :: current

        ! Count fields
        count    = 1
        in_quotes = .false.
        do i = 1, len_trim(line)
            if (line(i:i) == '"') in_quotes = .not. in_quotes
            if (line(i:i) == ',' .and. .not. in_quotes) count = count + 1
        end do

        allocate(fields(count))
        n_fields  = count
        count     = 1
        start     = 1
        in_quotes = .false.
        current   = ' '

        do i = 1, len_trim(line)
            if (line(i:i) == '"') then
                in_quotes = .not. in_quotes
            else if (line(i:i) == ',' .and. .not. in_quotes) then
                fields(count) = adjustl(current(1:i-start))
                count   = count + 1
                start   = i + 1
                current = ' '
            else
                current(i-start+1:i-start+1) = line(i:i)
            end if
        end do
        fields(count) = adjustl(current(1:len_trim(line)-start+1))
    end subroutine parse_csv_line

    pure function word_count(s) result(n)
        character(len=*), intent(in) :: s
        integer :: n, i
        logical :: in_word
        n = 0
        in_word = .false.
        do i = 1, len_trim(s)
            if (s(i:i) /= ' ' .and. s(i:i) /= char(9)) then
                if (.not. in_word) then
                    n = n + 1
                    in_word = .true.
                end if
            else
                in_word = .false.
            end if
        end do
    end function word_count

end module string_utils


program string_processing
    use string_utils
    implicit none

    character(len=256) :: s, t
    character(len=256), allocatable :: tokens(:), fields(:)
    integer :: n, i
    character(len=*), parameter :: csv_line = &
        '"Smith, John",42,"New York","Engineer, Senior",95000.50'

    print *, "=== String Processing Demo ==="
    print *, ""

    ! Basic operations
    s = "  Hello, World! This is Fortran 2018.  "
    print "(a, a)", "Original:   [", trim(s)//"]"
    print "(a, a)", "Upper:      [", trim(str_upper(s))//"]"
    print "(a, a)", "Lower:      [", trim(str_lower(s))//"]"
    print "(a, a)", "Trimmed:    [", trim(adjustl(s))//"]"
    print "(a, a)", "Reversed:   [", trim(reverse_str(adjustl(trim(s))))//"]"
    print "(a, i0)", "Word count: ", word_count(trim(s))
    print *, ""

    ! Search and replace
    s = "the quick brown fox jumps over the lazy dog"
    print "(a, a)", "Original: ", trim(s)
    print "(a, i0)", "Count 'the': ", count_occurrences(s, "the")
    t = replace_str(s, "the", "a")
    print "(a, a)", "Replace 'the'->'a': ", trim(t)
    print *, ""

    ! Predicates
    print *, "=== Type detection ==="
    do i = 1, 6
        select case (i)
        case (1); s = "42"
        case (2); s = "-17"
        case (3); s = "3.14159"
        case (4); s = "1.5e-3"
        case (5); s = "hello"
        case (6); s = "123abc"
        end select
        print "(a12, a, l1, a, l1)", trim(s), &
            "  is_integer=", is_integer(s), &
            "  is_real=", is_real(s)
    end do
    print *, ""

    ! Splitting
    s = "alpha:beta:gamma:delta:epsilon"
    print "(a, a)", "Split '", trim(s)//"' by ':'"
    tokens = str_split(trim(s), ":")
    n = size(tokens)
    do i = 1, n
        print "(a, i0, a, a)", "  Token ", i, ": ", trim(tokens(i))
    end do
    print "(a, a)", "Rejoined: ", trim(array_join(tokens, " | "))
    deallocate(tokens)
    print *, ""

    ! CSV parsing
    print *, "=== CSV Parsing ==="
    print "(a, a)", "Line: ", trim(csv_line)
    fields = str_getcsv(csv_line)
    n = size(fields)
    print "(a, i0, a)", "Parsed ", n, " fields:"
    do i = 1, n
        print "(a, i0, a, a)", "  [", i, "] ", trim(fields(i))
    end do
    deallocate(fields)
    print *, ""

    ! Internal file I/O (write/read to string)
    print *, "=== Internal File I/O ==="
    write(s, "(a, f8.3, a, i0, a)") "Pi = ", 4.0*atan(1.0), ", N = ", 42, " items"
    print "(a, a)", "Formatted string: ", trim(s)

    ! Read back from string
    block
        real :: pi_val
        integer :: n_val
        character(len=10) :: dummy
        read(s, *) dummy, pi_val, dummy, n_val
        print "(a, f8.3)", "Read back pi = ", pi_val
        print "(a, i0)",   "Read back N  = ", n_val
    end block

contains

    pure function str_upper(s) result(u)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: u
        integer :: i
        do i = 1, len(s)
            u(i:i) = to_upper(s(i:i))
        end do
    end function str_upper

    pure function str_lower(s) result(l)
        character(len=*), intent(in) :: s
        character(len=len(s)) :: l
        integer :: i
        do i = 1, len(s)
            l(i:i) = to_lower(s(i:i))
        end do
    end function str_lower

end program string_processing
