(function (window, undefined) {
	var apiCurrency = 'https://api.coingecko.com/api/v3/coins/markets?vs_currency=eur&order=market_cap_desc';

	function loadCurrency(callback) {
		var httpRequest;
		if (window.XMLHttpRequest) { // Mozilla, Safari, ...
			httpRequest = new XMLHttpRequest();
		} else if (window.ActiveXObject) { // IE
			try {
				httpRequest = new ActiveXObject("Msxml2.XMLHTTP");
			} catch (e) {
				try {
					httpRequest = new ActiveXObject("Microsoft.XMLHTTP");
				} catch (e) {
				}
			}
		}

		httpRequest.onreadystatechange = function () {
			if (httpRequest.readyState === 4) {
				callback(httpRequest.status === 200 ? httpRequest.responseText : null);
			}
		};
		httpRequest.open('GET', apiCurrency, true);
		httpRequest.send();
	}

	window.Asc.plugin.init = function () {
		this.button(-1);
	};

	window.Asc.plugin.button = function (id) {
		var t = this;
		loadCurrency(function (value) {
			var command = '';
			if (value) {
				try {
					var rates = JSON.parse(value);
					command += 'var oSheet = Api.GetActiveSheet();';
					command += 'var active = oSheet.GetActiveCell();';
					command += 'var row = active.GetRow();';
					command += 'var col = active.GetCol();';
					attributes = [
						'name',
						'low_24h',
						'current_price',
						'high_24h',
						'ath'
					];
					for (var i = 0; i < attributes.length; i++) {
						command += 'oSheet.GetRangeByNumber(row, col + ' + i + ').SetValue("' + attributes[i] + '");';
					}
					for (var j = 0; j < rates.length; j++) {
						for (var i = 0; i < attributes.length; i++) {
							command += 'oSheet.GetRangeByNumber(row + ' + (j+1) + ', col + ' + i + ').SetValue("' + rates[j][attributes[i]] + '");';
						}
					}

					window.Asc.plugin.info.recalculate = true;
				} catch (e) {
				}
			}
			console.log(command);
			t.executeCommand('close', command);
		});
	};
})(window, undefined);
