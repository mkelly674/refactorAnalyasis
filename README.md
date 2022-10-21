# refactorAnalyasis
## The Purpose

We are doing this for a friend name Steve who is starting to give stock advice for his career in finance after graduating. His first client was his parents and they wanted to invest in some green stocks. We refactor the workbook to see if it was faster.

## The Results

### Runtime

The runtime for the refactored code ran much quicker than the original workbook. As you can see in the image here:

![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/114030563/197107490-9feda22e-6936-48c3-a0ba-f6f354228c37.png)


The picture above is from the original and the picture below is the refactor.


![VBA_Challenge_2018_refactor](https://user-images.githubusercontent.com/114030563/197107506-744c41ab-e52e-4efb-bf88-7e753e6e187b.png)

As you can see, it ran almost 8 times faster with the refactor.

### The difference in the code

The thing that caused this lag was the "for" loop at the beginning that spanned the entire code in the original. The picture below is from the original code:
![original](https://user-images.githubusercontent.com/114030563/197283366-6882e3d6-d14e-4d07-98ff-6c7f26618f8f.png)

The refactored code had the "for" loops seperated and not nested. Here is what that looked like:
![refactor](https://user-images.githubusercontent.com/114030563/197283799-07fdbc49-b7e6-4d79-995e-4b0bbe55421d.png)

### The Cause in Lag

The reason why this nesting of "for" loops caused it is because it was running the code multiple times over instead of looping what it needed to loop. This also could cause runtime errors.



## The Conclusion

With everything that has been said, refactoring code can increase speed by eight times faster. If there is a lot of work that the computer has to be done, then the eight times difference will expand even more with better code. It can be even a difference between running the code and the computer overheating and not completing the task. With all that is said, sometimes refactoring it might make no sense because of how little the performace improved vs the amount of hours put into it that would cost more for a business. These apply here because it only improve the time in tenths of seconds which only can be seen as a little bit of lag.

