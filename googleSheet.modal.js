const mongoose = require('mongoose');

const spreadSheetSchema = mongoose.Schema(
  {
  url:{
    type: String ,
    unique: true,
    required: true,
  },
  date:{
    type: String,
    required: true,
  },
  time:{
    type:String,
    required : true
  }
  },
  {
    timestamps: true,
  }
);
/**
 * @typedef spreadSheet
 */
const spreadSheet = mongoose.model('spreadSheet', spreadSheetSchema);

module.exports = { 
    spreadSheet 
};