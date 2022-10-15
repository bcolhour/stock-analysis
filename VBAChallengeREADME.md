# Stock Analysis Refactored

## Overview
The client (Steve) wants a tool to analyze the entire stock market over multiple years. The current tool that was provided to the client was designed to analyze 12 stocks over 2 years. While this will work, it will be cumbersome and slow to analyze large volumes of data. The prior code for the original tool was analyzed and refactored to allow a more efficient processing of the data in a more timely manner. 

## Results
### Refactoring
The loop and variant coding was refactored to set 4 different arrays defined by a variant. A ticker index was also created. This allows for quicker execution by simplifying the steps using variants and removing unnecessary code. Conditional expressions were simplified to improve execution without affecting functionality.

![RefactoredCode](https://user-images.githubusercontent.com/114044192/195960844-8a5e7ba4-f80d-4e24-9f73-fcd5fe13fc22.PNG)
![RefactoredCode1](https://user-images.githubusercontent.com/114044192/195960856-5f844687-6cd2-4a0a-aa48-3eb7d17eec56.PNG)

### Outcome
The refactored code ran much quicker than the original tool coding. 

|Year | Original Code | Refactored Code |
|:---:|:-------------:|:---------------:|
|2017 | 1.585938 sec  |  0.265625 sec   |
|2018 | 1.570313 sec  |  0.2734375 sec  |

## Summary <sub>1</sub> 
### Refactoring Advantages
  - Correction of "Code Smells" and "Spaghetti Code" makes code more readable and functional
  - Removing duplications allows code to run smoother and allows for less chance of error
  - Allows adding new functionality to existing code without breaking existing code or making it difficult to execute
  - Improved legibility facilitates comprehension and code maintenance. 
  
### Refactoring Disadvantages
  - Efforts often not visible to end user
  - No precise guidelines of "neat code"
  - Possible introduction of new bugs or errors
  - Difficult and possible large effort required for larger teams to successfully refactor

### Refactoring VBA Script
Refactoring the VBA script allowed it to run much faster and cleaner without changing the desired functionality. As a new coder, this was difficult for me as there are no clear guidelines or rules as to what is the better code. Obviously removing nested loops, simplifying code and removing any duplication is optimal. However, I did introduce new errors that I had to debug at each step. 

#### Citation
<sub>1</sub> [Refactoring: How to improve source code](https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/ "Refactoring: How to improve source code")
