function colLetter = xlcolumnletter(colIndex)
    if colIndex <= 26
        colLetter = char('A' + colIndex - 1);
    else
        div = fix((colIndex - 1) / 26);
        rem = mod(colIndex - 1, 26);
        colLetter = [char('A' + div - 1), char('A' + rem)];
    end
end
