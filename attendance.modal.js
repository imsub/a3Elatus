const mongoose = require('mongoose');

const attendanceSchema = mongoose.Schema(
  {
    username: {
    type: String,
    unique: true,
    required: true,
  },
  globalName: {
    type: String,
    required: true,
  },
  nickName: {
    type: String,
    required: false,
  },
  userId: {
    type: String,
    unique: true,
    required: true,
  },
  attendance:{
    type: String,
    default: "present",
  },
  shareGiven :{
    type:String,
    default: "No"
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
 * @typedef Attendance
 */
const Attendance = mongoose.model('attendance', attendanceSchema);

module.exports = { 
    Attendance 
};