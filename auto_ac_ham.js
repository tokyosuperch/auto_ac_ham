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
var linearr = str.split('\r\n');

var accent_length = 60;
var hammering_length = 240;
var division = 480;
var time_sig = 4;

var position = {};

for (var lnum = 0; lnum < linearr.length; lnum++) {
    if (linearr[lnum].split("|").length < 4) {
        WScript.Echo(linearr[lnum]);
        continue;
    }
    if (linearr[lnum].match("MThd")) {
        WScript.Echo(file_init(linearr, lnum));
        continue;
    }
    if (linearr[lnum].match("Off Note")) {
        WScript.Echo(off_note(linearr, lnum));
        continue;
	}
    WScript.Echo(linearr[lnum]);
};

function file_init(linearr, lnum) {
    var pipe = linearr[lnum].split("|");
    return(pipe.join("|"));
};

function off_note(linearr, off_num) {
    var res_text = "";
    var timestr_off = event_time(linearr, off_num, null, null, null);
    var timestr_on;
    var on_num;
    for (on_num = off_num; on_num >= 0; on_num--) {
        if (linearr[on_num].match("On Note")) {
            timestr_on = event_time(linearr, on_num, null, null, null);
            break;
		}
	};
    if (note_length(timestr_on, timestr_off) >= hammering_length * 4) {
        res_text += event_time(linearr, on_num, Number(timestr_on.split(":")[0]), Number(timestr_on.split(":")[1]), Number(timestr_on.split(":")[2])) + "|Pitch Wheel | chan= 1   | bend=-1396\r\n";
        res_text += event_time_add(linearr, on_num, Number(timestr_on.split(":")[0]), Number(timestr_on.split(":")[1]), Number(timestr_on.split(":")[2]), hammering_length) + "|Pitch Wheel | chan= 1   | bend=0\r\n";
    }
    res_text += linearr[off_num];
    return res_text;
    // Console.WriteLine("\t" + String.Format("{0,3}", currenttick + 60) +" |Pitch Wheel | chan= 1   | bend=4096");
    // Console.WriteLine("\t" + String.Format("{0,3}", currenttick + 120) + " |Pitch Wheel | chan= 1   | bend=0");
};

function event_time_add(linearr, lnum, mea, beat, tick, length) {
    var new_tick = tick + length;
    var new_beat = beat;
    var new_mea = mea;
    if (new_tick >= division) {
        new_tick -= division;
        new_beat += 1;
	}
    if (new_beat > time_sig) {
        new_beat -= 4;
        new_mea += 1;
	}
    return event_time(linearr, lnum, new_mea, new_beat, new_tick);
};

function event_time(linearr, lnum, mea, beat, tick) {
    var pipe = linearr[lnum].split("|");
    var timearr = pipe[0].split(":");
    
    if (timearr.length == 1) {
        if (timearr[0].match(/[0-9]/)) {
            if (tick == null) tick = Number(timearr[0]);
		}
        return event_time(linearr, lnum-1, mea, beat, tick);
	} else if (timearr.length == 2) {
        if (beat == null) beat = Number(timearr[0]);
        if (tick == null) tick = Number(timearr[1]);
        return event_time(linearr, lnum-1, mea, beat, tick);
	} else if (timearr.length == 3) {
        if (mea == null) mea = Number(timearr[0]);
        if (beat == null) beat = Number(timearr[1]);
        if (tick == null) tick = Number(timearr[2]);
        return ("    " + mea).substr(-4) + ":" + ("  " + beat ).substr(-2) + ":" + ("   " + tick ).substr(-3) + " ";
	} else {
        return null;
	}
};

function note_length(timestr_on, timestr_off) {
    var off_arr = timestr_off.split(":");
    var on_arr = timestr_on.split(":");
    for (var n = 0; n < off_arr.length; n++) {
        off_arr[n] = Number(off_arr[n]);
	};
    for (var n = 0; n < on_arr.length; n++) {
        on_arr[n] = Number(on_arr[n]);
	};

    var totaltick = {};
    totaltick.off = ((off_arr[0] - 1) * time_sig * division) + ((off_arr[1] - 1) * division) + off_arr[2];
    totaltick.on = ((on_arr[0] - 1) * time_sig * division) + ((on_arr[1] - 1) * division) + on_arr[2];
    return totaltick.off - totaltick.on;
};