const mongoose = require('mongoose');

const JSONVarSchema = mongoose.Schema({
    customId: {
        type: Number,
        default: 1
    },
    json: {
        type: String,
        default: `{
            "gss" : {
                "sheet1": {
                    "arco": "D",
                    "agrobox": "F",
                    "up": "G"
                },
                "sheet2": {
                    "arco": "C",
                    "up": "I",
                    "pwc": "G",
                    "willemen": "E"
                },
                "sheet3": {
                    "arco": "C",
                    "up": "E",
                    "upc": "G",
                    "agrobox": "H",
                    "jefacture": "N",
                    "pwc": "I",
                    "spotcheck": "D",
                    "willemen": "K"
                },
                "sheet4": {
                    "agrobox": "E",
                    "upc": "F",
                    "up": "G"
                },
                "sheet5": {
                    "up": "E",
                    "upc": "F",
                    "agrobox": "H"
                },
                "sheet6": {
                    "arco": "D",
                    "pwc": "E"
                },
                "sheet7": {
                    "agrobox": "C",
                    "up": "D",
                    "spotcheck": "F"
                },
                "sheet9": {
                    "agrobox": "C",
                    "up": "D"
                }
            },
            "transpprice": {
                "day": "2000",
                "night": "1000"
            },
            "repas": "3500",
            "spotckeck_mult": "100",
            "willemen": {
                "ai": "30",
                "id": "50",
                "limosa": "100"
            }
        }`
    }
})

module.exports = mongoose.model('jsonVar', JSONVarSchema);