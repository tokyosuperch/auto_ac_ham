function file_open(file, encode){
    var stream = WScript.CreateObject('ADODB.Stream');
    stream.charset = encode || "Shift-JIS";
    stream.Open();
    stream.loadFromFile(file);
    var contents = stream.ReadText();
    stream.close();

    return contents;
}

// substr()‚Ìpolyfill
// only run when the substr() function is broken
if ('ab'.substr(-1) != 'b') {
  /**
   *  Get the substring of a string
   *  @param  {integer}  start   where to start the substring
   *  @param  {integer}  length  how many characters to return
   *  @return {string}
   */
  String.prototype.substr = function(substr) {
    return function(start, length) {
      // call the original method
      return substr.call(this,
      	// did we get a negative start, calculate how much it is from the beginning of the string
        // adjust the start parameter for negative value
        start < 0 ? this.length + start : start,
        length)
    }
  }(String.prototype.substr);
}
// polyfillI‚í‚è

var oArgs = WScript.Arguments;
var arg = oArgs.Unnamed(0);
// WScript.Echo();
var str = file_open(arg, "Shift-JIS");
var line = str.split('\r\n');

var position = {};

for (var l = 0; l < line.length; l++) {
    if (line[l].split("|").length < 4) {
        WScript.Echo(line[l]);
        continue;
    }
    if (line[l].match("MThd")) {
        WScript.Echo(file_init(line, l));
        continue;
    }
    if (line[l].match("On Note")) {
        WScript.Echo(on_note(line, l));
        continue;
	}
};

function file_init(line, l) {
    var pipe = line[l].split("|");
    return(pipe.join("|"));
};

function on_note(line, l) {
    return note_time(line, l, null, null, null);
    Console.WriteLine("\t" + String.Format("{0,3}", currenttick + 60) +" |Pitch Wheel | chan= 1   | bend=4096");
    Console.WriteLine("\t" + String.Format("{0,3}", currenttick + 120) + " |Pitch Wheel | chan= 1   | bend=0");
};

function note_time(line, l, mea, beat, tick) {
    var pipe = line[l].split("|");
    var timearr = pipe[0].split(":");
    
    if (timearr.length == 1) {
        if (timearr[0].match(/0-9/)) {
            if (tick == null) tick = Number(timearr[0]);
		}
        return note_time(line, l-1, mea, beat, tick);
	} else if (timearr.length == 2) {
        if (beat == null) beat = Number(timearr[0]);
        if (tick == null) tick = Number(timearr[1]);
        return note_time(line, l-1, mea, beat, tick);
	} else if (timearr.length == 3) {
        mea = Number(timearr[0]);
        if (beat == null) beat = Number(timearr[1]);
        if (tick == null) tick = Number(timearr[2]);
        return ("    " + mea).substr(-4) + ":" + ("  " + beat ).substr(-2) + ":" + ("   " + tick ).substr(-3) + " ";
	} else {
        return null;
	}
};