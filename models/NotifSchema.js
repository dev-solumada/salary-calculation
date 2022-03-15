const mongoose = require('mongoose');

const NotifSchema = mongoose.Schema({
    category: {
        type: String,
        default: 'none'
    },
    description: {
        type: String,
        default: 'none'
    },
    seen: {
        type: Boolean,
        default: false
    },
    creation: {
        type: Date,
        default: new Date()
    },
})

module.exports = mongoose.model('notifschema', NotifSchema);