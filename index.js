var util = require('util'),
    queue = require('./queue'),
    axios = require('axios'),
    winston = require('winston'),
    webColors = {
        "black": "000",
        "red": "f00",
        "green": "0f0",
        "yellow": "ff0",
        "blue": "00f",
        "magenta": "f0f",
        "cyan": "0ff",
        "white": "fff",
        "gray": "808080",
        "grey": "808080"
    };
const Transport = require('winston-transport');

class Office365ConnectorHook extends Transport {
    constructor(options) {
        super(options);
        this.name = options.name || 'office365connectorhook';
        this.level = options.level || 'silly';
        this.hookUrl = options.hookUrl || null;
        this.webColors = options.colors || {};
        this.prependLevel = options.prependLevel === undefined ? true : options.prependLevel;
        this.appendMeta = options.appendMeta === undefined ? true : options.appendMeta;
        this.formatter = options.formatter || null;
    }

    log(info, callback) {
        let color = winston.config.syslog.colors[info.level],
            themeColor = this.webColors[info.level] || webColors[color],
            title = "";
        let message = '';
        if (this.prependLevel) {
            message += '[' + info.level + ']\n';
        }
        message += info.message;

        if (this.appendMeta && info.meta) {
            title = info.meta.title || "";
            delete info.meta["title"];
            var props = Object.getOwnPropertyNames(info.meta);
            if (props.length) {
                // http://stackoverflow.com/questions/18391212/is-it-not-possible-to-stringify-an-error-using-json-stringify#comment57014279_26199752
                message += ' ```' + JSON.stringify(info.meta, props, 2) + '```';
            }
        }

        if (typeof this.formatter === 'function') {
            message = this.formatter({
                level: info.level,
                message: message,
                meta: info.meta
            });
        }

        let payload = {
            title: title,
            text: message,
            themeColor: themeColor
        };

        q.push(deferredExecute(this.hookUrl, payload, callback));
    }
}

var q = queue({ concurrency: 1, autostart: true });

ensureCallback = function (cb1, cb2) {
    return function () {
        try {
            cb1.apply(null, arguments);
        }
        finally {
            cb2();
        }
    }
}

deferredExecute = function (url, payload, callback) {
    return function (cb) {
        var safeCallback = ensureCallback(callback, cb);
        axios.post(url, payload)
            .then(response => {
            if (response.status === 200) {
                safeCallback(null, true);
            } else {
                safeCallback('Server responded with ' + response.status);
            }
            })
            .catch(error => {
            safeCallback(error);
            });
    }
}

module.exports = Office365ConnectorHook;