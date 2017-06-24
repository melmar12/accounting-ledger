// Checks if everything is ok

function validate()
{
    sumof=parseFloat(0.0);
    err_count = 0;

    s0=0;
    crlf = String.fromCharCode(10, 13);
    // msgwin is big textbox on bottom, displays error messages
    document.a3form.msgwin.value = " ";

    // Begin checking Fiscal Month
    if (document.a3form.fiscal.value.length != 0)
    // Is there a character
    {
        if (isNaN(document.a3form.fiscal.value))
        // Is the character a number
        {
            err_count=err_count+1;
            document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+ err_count.toString() + ". Fiscal Month Must be Numeric";
        }
        else
        {
            if (parseInt(document.a3form.fiscal.value) < 1 || parseInt(document.a3form.fiscal.value) > 12)
            // Is the character a valid fiscal month
            {
                err_count=err_count+1
                document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+ err_count.toString()+ ". Fiscal Month must be 1,2,3,...,12";
            }
        }
    }
    else
    // Textbox is empty
    {
        err_count=err_count+1;
        document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+ err_count.toString() + ". Must Enter a Fiscal Month";
    }

    // Begin checking Source Reference Number
    if (document.a3form.sourceref.value.length != 4)
    // Are there 4 digits in the source reference number
    {
        err_count=err_count+1;
        document.a3form.msgwin.value = document.a3form.msgwin.value + crlf + err_count.toString() + ". Source Ref. Number Must be 4 digits";
    }
    else
    {
        if (isNaN(document.a3form.sourceref.value))
        // Is the source reference number a number
        {
            err_count=err_count+1;
            document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+err_count.toString() + ". Source Ref. Number must be numbers ONLY";
        }
    }

    document.a3form.numvalid.value = 0;
    // End check top row
    // Grab journal entry
    for (entrynum = 1; entrynum < 8;entrynum++)
    {
        s0=0;
        // s0 helps keep track of error
        point = entrynum * 6 - 4;
        la1 = document.a3form.elements[point + 0].value.length;
        // Length of major
        la2 = document.a3form.elements[point + 1].value.length;
        // Length of minor
        la3 = document.a3form.elements[point + 2].value.length;
        // Length of sub1
        la4 = document.a3form.elements[point + 3].value.length;
        // Length of sub2
        la5 = document.a3form.elements[point + 4].value.length;
        // Length of transaction amount
        tbytes=la1+la2+la3+la4+la5

        if (tbytes!=0) // there is something in accts and transaction amt
        {
            e_str = entrynum.toString();
            for (glc = 0; glc < 4; glc++)
                // Run through major, minor, sub1, sub2
            {
                fred = document.a3form.elements[point + glc].value;
                // Fred takes texbox value
                if (fred != 0)
                {
                    s0=1;
                    if (isNaN(fred))
                        // If true, spit out error
                    {
                        err_count=err_count+1;
                        if (glc == 0)
                            document.a3form.msgwin.value = document.a3form.msgwin.value + crlf + err_count.toString() + ". Major on Journal Entry " + e_str + " must be numeric";
                        if (glc == 1)
                            document.a3form.msgwin.value = document.a3form.msgwin.value + crlf + err_count.toString() + ". Minor on Journal Entry " + e_str + " must be numeric";
                        if (glc == 2)
                            doocument.a3form.msgwin.value = document.a3form.msgwin.value + crlf + err_count.toString() + ". Sub1 on Journal Entry " + e_str + " must be numeric";
                        if (glc == 3)
                            document.a3form.msgwin.value = document.a3form.msgwin.value + crlf + err_count.toString() + ". Sub2 on Journal Entry " + e_str + " must be numeric";
                    }
                    else
                    {
                        document.a3form.elements[point+glc].value= makefour(fred);
                        fred= document.a3form.elements[point+glc].value
                    }
                }
                else
                {
                    if (s0 == 1)
                        document.a3form.elements[point + glc].value = "0000";
                }
            }
            fred = document.a3form.elements[point + 4].value;
            // Take in value of transaction amount
            if  (s0  > 0 && fred.length !=0 )
            {
                if (isNaN(fred))
                {
                    err_count=err_count+1;
                    document.a3form.msgwin.value = document.a3form.msgwin.value + crlf +err_count.toString() + ". Transaction Amount" + " on entry " + e_str + " must be numeric";
                }
                else
                    // If transaction amount has value, add it to sumof variable
                {
                    document.a3form.elements[point+4].value= roundit(fred);
                    sumof=sumof + parseFloat(document.a3form.elements[point+4].value);
                    s0=2;
                }
            }
            else
            {
                if (s0>0 && fred.length==0)
                // If no input in transaction amount, spit out error
                {
                    err_count=err_count+1;
                    document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+ err_count.toString()+". Valid Account " + "on entry " + e_str+ " must have a Transaction Amount";
                }
                else
                {
                    if (s0== 0 && fred.length !=0)
                    {
                        err_count=err_count+1;
                        document.a3form.msgwin.value = document.a3form.msgwin.value + crlf+err_count.toString() + ". Trans. Amt " +"on entry " + e_str +" must have a Valid GL Account Number";
                    }
                }
            }
    }
        if (s0 == 2)
            document.a3form.numvalid.value = parseInt(document.a3form.numvalid.value) + 1;
            // Hidden object that holds number that of inputs that passed. Counts correct journal entries
    }
    // end of loop
    if (sumof > -0.000001 && sumof < 0.000001)
    // if sum is a super small number close to 0, it is 0
        sumof =0.0;
    document.a3form.sumofamt.value = sumof.toFixed(2);
    // set to 2 decimals
    if (sumof !=0.0)
    {
        // inputs must = 0
        err_count=err_count+1;
        document.a3form.msgwin.value = document.a3form.msgwin.value + crlf +err_count.toString()+ ". Sum of Transaction Amounts not zero";
    }
    if (parseInt(a3form.numvalid.value) < 2)
    {
        // if only < 2 passed the input tests
        err_count=err_count+1;
        document.a3form.msgwin.value = document.a3form.msgwin.value + crlf +err_count.toString()+". " + document.a3form.numvalid.value.toString() + " valid entries. Must have 2 or more valid entries";
    }

    if (sumof == 0 && err_count == 0)
    {
        // If journal entries add up to 0, and no errors were found
        document.a3form.msgwin.value= "Finished Validation: Found: "+ a3form.numvalid.value.toString()+"  valid entries; No Errors."+crlf +"Submitting the Journal Voucher for Processing";
        document.a3form.submit();
    }
    document.a3form.msgwin.value = "Finished Validation: Found: "+ document.a3form.numvalid.value.toString() +" valid entries; "+ err_count.toString() + " Error(s) Detected:" + document.a3form.msgwin.value;
}

function makefour(afield)
{
    make4val="";
    length_of_afield=afield.length;
    if (length_of_afield == 4)
        make4val = afield;
    if (length_of_afield == 3)
        make4val = "0" + afield;
    if (length_of_afield == 2)
        make4val = "00" + afield;
    if (length_of_afield == 1)
        make4val = "000" + afield;

    return make4val;
}

function roundit(afield)
{
    bfield=afield*100.0;
    bfield=parseInt(bfield);
    bfield=parseFloat(bfield)/100.0;
    rounditval=bfield.toFixed(2);
    return rounditval;
}
