function httpGet(url) {
	fetch(url, {
		headers: {
			'Content-Type': 'application/json; charset=UTF-8'
		},
		method: 'GET'
	}).then(response => response.text())
		.catch(e => null)
}

function nextPage() {
	httpGet('/next');
}

function previousPage() {
	httpGet('/previous');
}

function slideNumber(slideNumberValue) {
	httpGet('/slideNumber' + slideNumberValue);
}

document.addEventListener('swiped-left', function (e) {
	nextSlideAll();
})
document.addEventListener('swiped-right', function (e) {
	previousSlideAll();
})