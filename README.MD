This project is designed to ease the process of keeping track of my
finances. More specifically, I'd like to keep track of month to
month changes of my bank account.

The input is a well-formatted excel file that holds the amounts in my various
accounts. The output is an excel file with a summary sheet and sheets for each
individual tabulation.

Current Status:
  Program takes an input file and a desired output file. It handles whether or not the output file
  has been used yet or not.

  The input file should have well-formatted inputs inside. It should be formatted as follows:

  Bank Account | Totals
  Checking     | ___
  Saving       | ___
  Money Market | ___

  Additionally, further down in the columns, you can insert an investment block of information.

  It will look as follows:
  Investment           | Values
  Date Invested        | ___
  Date To Be Returned  | ___
  Principal            | ___
  Interest Rate        | ___
  Length of Investment | ___
  New Money            | ___

  The crucial aspects of naming are in the "Bank Account" cell and the "Investment" cell. These must exactly match as described.
  Also, the values obviously must be inserted where the blank spaces are. However, while the values should be the same meaning,
  the other names actually can be tuned to your liking.
  (Hint: "New Money" helps keep track of the total amount of money invested, so this should only be non-zero when money is being
  injected into the investment)
  (Hint: If you only want to insert an investment then delete the word "Bank Account", if you only want to insert a Bank Account track,
  delete the word "Investment", these are the names that the program looks for, marking that it has an account or investment to insert,
  by deleting these words, it won't attempt to insert or account for them)

  The output file will have a summary sheet along with an archive of all the bank account tracking
  The summary sheet will have the current total along with various deltas over certain periods of time.
  There will also be archives for both the investments and tracking of my bank account. They will look as described below.

  The summary sheet will have a format like the following:
  Bank Account Totals         | Current | Delta Over 1 Month     | Delta Over 3 Months | Delta Over 6 Months | Delta Over 9 Months | Delta Over 1 Year
  Checking                    |         |                        |                     |                     |                     |  
  Savings                     |         |                        |                     |                     |                     |
  Money Market                |         |                        |                     |                     |                     |
  Total                       |         |                        |                     |                     |                     |
                              |         |                        |                     |                     |                     |
  Investment Totals           |         | Most Recent Investment |                     |                     |                     |
  Total Invested              |         | Date Invested          |                     |                     |                     |
  Total Profit                |         | Date To Be Returned    |                     |                     |                     |
  # of Re-Investments         |         | Principal              |                     |                     |                     |
  Time since beginning (days) |         | Rate                   |                     |                     |                     |
                              |         | Length (days)          |                     |                     |                     |
                              |         | Profit                 |                     |                     |                     |

  The Bank Archive will look like the following, with values extending downwards:
  Bank Archive | Checking | Savings | Money Market | Total

  The Investment Archive will look like the following, also with values extending downwards:
  Date Invested | Date Returned | Principal | Rate | Length of Investment (days) | Profit


Future Status:
  Program will also handle individual investments/loans, student debt, etc. All in a similar fashion
