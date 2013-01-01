/* summarize the data and store */
/* the output in an output data set */
proc means data=&data &stat noprint;
	var &measure;
	class &report;
	output out=summary &stat=&measure /levels;
run;

/* store the value of the measure for ALL rows and 
/* the row count into a macro variable for use  */
/* later in the report */
proc sql noprint;
select &measure,_FREQ_ into :overall,:numobs
from summary where _TYPE_=0;
quit;

/* sort the results so that we get the TOP values */
/* rising to the top of the data set */
proc sort data=work.summary out=work.topn;
  where _type_>0;
  by descending &measure;
run;

/* Pass through the data and output the first N */
/* values for each category */
data topn;
  set topn;
  by descending &measure;
  rank+1;
  if rank le &n then output;
run;

/* Create a report listing for the top values in each category */
title "&stat of &measure for all values of &report: &overall"; 
title2 "&numobs total rows";
proc report data=topn;
	columns rank &report &measure;
	define rank /display;
run;
quit; 

title;
/* Create a simple bar graph for the data to show the rankings */
/* and relative values */
goptions xpixels=600 ypixels=400;
proc gchart data=topn
;
	hbar &report / 
		sumvar=&measure
		descending
		nozero
		clipref
		frame	
		discrete
		type=&stat
;
run;
quit;