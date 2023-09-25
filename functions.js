const signatureSetting = 'calendarSignature';

let existingBody = '';
let requestHeader = undefined;

Office.onReady(function(info) {
	// note: used in both the add signature button and the standalone taskpane, so we need to check objects exist
	if (info.host === Office.HostType.Outlook) {
		if (window.jQuery) {
			$(document).ready(function() {
				const textColours = [
					'000000', '262626', '595959', '8c8c8c', 'bfbfbf', 'd9d9d9', 'e9e9e9', 'f5f5f5', 'fafafa', 'ffffff',
					'fbe9e6', 'fcede1', 'fcefd4', 'fcfbcf', 'e7f6d5', 'daf4f0', 'd9edfa', 'e0e8fa', 'ede1f8', 'f6e2ea',
					'ffa39e', 'ffbb96', 'ffd591', 'fffb8f', 'b7eb8f', '87e8de', '91d5ff', 'adc6ff', 'd3adf7', 'ffadd2',
					'ff4d4f', 'ff7a45', 'ffa940', 'ffec3d', '73d13d', '36cfc9', '40a9ff', '597ef7', '9254de', 'f759ab',
					'e13c39', 'e75f33', 'eb903a', 'f5db4d', '72c040', '59bfc0', '4290f7', '3658e2', '6a39c9', 'd84493',
					'cf1322', 'd4380d', 'd46b08', 'd4b106', '389e0d', '08979c', '096dd9', '1d39c4', '531dab', 'c41d7f',
					'820014', '871400', '873800', '614700', '135200', '00474f', '003a8c', '061178', '22075e', '780650'
				];

				const trumbowygElement = $('#signatureText');
				trumbowygElement.trumbowyg({
					btns: [
						['fontfamily', 'fontsize'],
						['bold', 'italic', 'underline'],
						['foreColor', 'backColor'],
						['link'],
						['viewHTML']
					],
					semantic: false,
					plugins: {
						colors: {
							foreColorList: textColours,
							allowCustomForeColor: true,
							backColorList: textColours,
							allowCustomBackColor: true,
						}
					},
					autogrow: false
				});

				if (Office.context.roamingSettings.get(signatureSetting)) {
					trumbowygElement.trumbowyg('html', 
								Base64.decode(Office.context.roamingSettings.get(signatureSetting)));
				}

				let height = 100; // to avoid button overlap at small sizes
				trumbowygElement.closest('.trumbowyg-box').css({minHeight: height});
				height -= $('.trumbowyg-button-pane').height() - 1; // button bar is 36, but can wrap; 1 is the border
				trumbowygElement.prev('.trumbowyg-editor').css({minHeight: height});
				trumbowygElement.css({minHeight: height}); // fixes the raw HTML editor

				$(window).on('resize', resizeEditor);
				resizeEditor();
			});
		}

		// elements are hidden by default (in case Office is not present)
		const signatureBox = document.getElementById('signatureText');
		if (signatureBox) {
			signatureBox.style.display = 'block';
		}
		const saveButton = document.getElementById('saveSignature');
		if (saveButton) {
			saveButton.style.display = 'block';
			saveButton.onclick = saveSignature;
		}
	}
});

function showTaskPaneMessage(messageText, autoHide) {
	if (window.jQuery) {
		$('#messageBox').text(messageText).show().delay(2500).queue(function (next) {
			if (autoHide) {
				$(this).hide();
			}
			next();
		});
	} else {
		console.log(messageText);
	}
}

function resizeEditor() {
	const saveButton = $('#saveSignature');
	const trumbowygElement = $('#signatureText');
	const closestBox = trumbowygElement.closest('.trumbowyg-box');

	// the first launch popup is shown in a different style, so needs a slight visual edit
	if (window.location.search.indexOf('?firstlaunch') !== -1) {
		$('body').css({padding: 0});
		saveButton.css({margin: 0, marginRight: 6});
	}

	// make the editor fill the available height
	let newHeight = saveButton.position().top - closestBox.position().top - 12; // 12 is padding between button and box
	closestBox.css({height: newHeight});
	newHeight = newHeight - $('.trumbowyg-button-pane').height() - 1; // button bar is 36, but can wrap; 1 is the border
	trumbowygElement.prev('.trumbowyg-editor').css({height: newHeight}); // the main editor box
	trumbowygElement.css({height: newHeight}); // the textarea (used for HTML editing)
}

function cleanIDAndClassValues(bodyContent) {
	// need to remove any id="[id]" (+ class="") because Outlook replaces it with id="x_[id]", which breaks detection
	return bodyContent.replaceAll(/(id|class)="x_(.*?)"/gi, '$1="$2"');
}

function saveSignature() {
	let hasSavedSignature = Office.context.roamingSettings.get(signatureSetting);

	// we need to remove any id="[id]" (+ class="") because Outlook replaces it with id="x_[id]", which breaks detection
	let signature;
	if (window.jQuery) {
		signature = cleanIDAndClassValues($('#signatureText').trumbowyg('html'));
	} else {
		signature = cleanIDAndClassValues(document.getElementById('signatureText').value);
	}

	showTaskPaneMessage('✎    Saving…', false);
	Office.context.roamingSettings.set(signatureSetting, Base64.encode(signature));
	Office.context.roamingSettings.saveAsync(function(asyncResult) {
		if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
			showTaskPaneMessage('✓    Saved', true);
			if (!hasSavedSignature) {
				if (typeof Office.context.ui.messageParent == 'function') {
					Office.context.ui.messageParent(true.toString()); // first time - close dialog and re-add signature
				}
			}
			// Office.addin.hide(); // not supported on the web
		} else {
			showTaskPaneMessage('✗    Save failed', false);
			console.error('saveSignature', asyncResult.status, ':', asyncResult.error.message);
		}
	});
}

function addSignature(firstTime) {
	// if triggered from the ribbon button (i.e., not a taskpane) we get an anonymous callback and notify once complete
	if (typeof firstTime === 'object') {
		try {
			if (typeof firstTime.completed === 'function') {
				requestHeader = firstTime;
			}
		} catch (error) {
		}
		firstTime = true;
	}

	if (!Office.context.roamingSettings.get(signatureSetting)) {
		console.log('addSignature: no saved signature found; opening save dialog');
		
		// we need a separate dialog box on first load because Office.addin.showAsTaskpane() isn't supported in Outlook
		Office.context.ui.displayDialogAsync(
			'https://simonrob.github.io/ocsa/taskpane.html?firstlaunch',
			{width: 60, height: 70, displayInIframe: true},
			function (asyncResult) {
				if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
					const dialog = asyncResult.value;
					dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (message) {
						dialog.close(); // the only message we ever send is "signature saved"; no need to check content

						// it would be better to call addSignature (with requestHeader) to add the signature straight
						// away, but there is no way (without a page refresh) to reload the roamingSettings object; if 
						// we instead call addSignature from the save callback, Office.context.mailbox is undefined...
						if (typeof requestHeader === 'object') {
							requestHeader.completed();
							requestHeader = undefined;
						}
					});
				} else {
					console.log('displayDialogAsync', asyncResult.status, ':', asyncResult.error.message);
				}
			}
		);
		return;
	}

	// append the signature to any existing body content
	// note: we need to do this because while there is a setSignatureAsync method, it doesn't work with calendar events 
	// ("The operation is not supported"): https://learn.microsoft.com/en-us/javascript/api/outlook/office.body?view=
	// outlook-js-preview#outlook-office-body-setsignatureasync-member(1)
	Office.context.mailbox.item.body.getAsync(
		Office.CoercionType.Html,
		function (asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				appendSignature(false, asyncResult.value + 
					Base64.decode(Office.context.roamingSettings.get(signatureSetting)));
				
				/*
				// note: removed this more complex approach as it needs more testing in desktop apps 
				if (firstTime) {
					// first step: get the existing body to later check whether it has a signature already
					existingBody = asyncResult.value;
					appendSignature(false, Base64.decode(Office.context.roamingSettings.get(signatureSetting)));
				} else {
					// third step: asyncResult.value is now the signature as re-formatted by Outlook
					appendSignature(false, cleanIDAndClassValues(existingBody.replace(asyncResult.value, '')) + 
						asyncResult.value);
				}
				*/
			} else {
				console.log('addSignature', asyncResult.status, ':', asyncResult.error.message);
			}
		}
	);
}

function appendSignature(firstTime, bodyContent) {
	Office.context.mailbox.item.body.setAsync(
		bodyContent,
		{ coercionType: Office.CoercionType.Html },
		function (asyncResult) {
			if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
				if (firstTime) {
					// second step: add just the signature, replacing the existing body
					addSignature(false);
				} else {
					// final step: hide the "working on your request" message
					if (typeof requestHeader === 'object') {
						requestHeader.completed();
						requestHeader = undefined;
					}
				}
			} else{
				console.log('appendSignature', asyncResult.status, ':', asyncResult.error.message);
			}
		}
	);
}


/**
 *
 *  UTF-8-safe Base64 encode / decode utility
 *  https://www.webtoolkit.info/javascript_base64.html
 *
 **/
var Base64 = {
	// private property
	_keyStr: 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=',

	// public method for encoding
	encode: function(input) {
			var output = '';
			var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
			var i = 0;
			input = Base64._utf8_encode(input);
			while (i < input.length) {
				chr1 = input.charCodeAt(i++);
				chr2 = input.charCodeAt(i++);
				chr3 = input.charCodeAt(i++);
				enc1 = chr1 >> 2;
				enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
				enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
				enc4 = chr3 & 63;
				if (isNaN(chr2)) {
					enc3 = enc4 = 64;
				} else if (isNaN(chr3)) {
					enc4 = 64;
				}
				output = output + this._keyStr.charAt(enc1) + this._keyStr.charAt(enc2) + 
							this._keyStr.charAt(enc3) + this._keyStr.charAt(enc4);
			} // Whend 
			return output;
		}, // End Function encode 

	// public method for decoding
	decode: function(input) {
			var output = '';
			var chr1, chr2, chr3;
			var enc1, enc2, enc3, enc4;
			var i = 0;
			input = input.replace(/[^A-Za-z0-9\+\/\=]/g, '');
			while (i < input.length) {
				enc1 = this._keyStr.indexOf(input.charAt(i++));
				enc2 = this._keyStr.indexOf(input.charAt(i++));
				enc3 = this._keyStr.indexOf(input.charAt(i++));
				enc4 = this._keyStr.indexOf(input.charAt(i++));
				chr1 = (enc1 << 2) | (enc2 >> 4);
				chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
				chr3 = ((enc3 & 3) << 6) | enc4;
				output = output + String.fromCharCode(chr1);
				if (enc3 != 64) {
					output = output + String.fromCharCode(chr2);
				}
				if (enc4 != 64) {
					output = output + String.fromCharCode(chr3);
				}
			} // Whend 
			output = Base64._utf8_decode(output);
			return output;
		}, // End Function decode 

	// private method for UTF-8 encoding
	_utf8_encode: function(string) {
			var utftext = '';
			string = string.replace(/\r\n/g, '\n');
			for (var n = 0; n < string.length; n++) {
				var c = string.charCodeAt(n);
				if (c < 128) {
					utftext += String.fromCharCode(c);
				} else if ((c > 127) && (c < 2048)) {
					utftext += String.fromCharCode((c >> 6) | 192);
					utftext += String.fromCharCode((c & 63) | 128);
				} else {
					utftext += String.fromCharCode((c >> 12) | 224);
					utftext += String.fromCharCode(((c >> 6) & 63) | 128);
					utftext += String.fromCharCode((c & 63) | 128);
				}
			} // Next n 
			return utftext;
		}, // End Function _utf8_encode 

	// private method for UTF-8 decoding
	_utf8_decode: function(utftext) {
		var string = '';
		var i = 0;
		var c, c1, c2, c3;
		c = c1 = c2 = 0;
		while (i < utftext.length) {
			c = utftext.charCodeAt(i);
			if (c < 128) {
				string += String.fromCharCode(c);
				i++;
			} else if ((c > 191) && (c < 224)) {
				c2 = utftext.charCodeAt(i + 1);
				string += String.fromCharCode(((c & 31) << 6) | (c2 & 63));
				i += 2;
			} else {
				c2 = utftext.charCodeAt(i + 1);
				c3 = utftext.charCodeAt(i + 2);
				string += String.fromCharCode(((c & 15) << 12) | ((c2 & 63) << 6) | (c3 & 63));
				i += 3;
			}
		} // Whend 
		return string;
	} // End Function _utf8_decode 
}
