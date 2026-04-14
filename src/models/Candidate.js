import mongoose from 'mongoose';
import { SUBJECT_KEYS } from '../constants/subjects.js';

const scoresSchemaDefinition = SUBJECT_KEYS.reduce((accumulator, key) => {
  accumulator[key] = {
    type: Number,
    min: 0,
    max: 300,
    default: null
  };
  return accumulator;
}, {});

const candidateSchema = new mongoose.Schema(
  {
    examNumber: {
      type: String,
      required: true,
      unique: true,
      trim: true,
      index: true
    },
    fullName: {
      type: String,
      required: true,
      trim: true
    },
    birthDate: {
      type: String,
      default: ''
    },
    className: {
      type: String,
      default: ''
    },
    scores: scoresSchemaDefinition
  },
  {
    timestamps: true,
    toJSON: {
      virtuals: true,
      transform: (_, returnedObject) => {
        delete returnedObject.__v;
        return returnedObject;
      }
    }
  }
);

candidateSchema.virtual('totalScore').get(function totalScore() {
  return SUBJECT_KEYS.reduce((total, subjectKey) => {
    const value = this.scores?.[subjectKey];
    return typeof value === 'number' ? total + value : total;
  }, 0);
});

const Candidate = mongoose.model('Candidate', candidateSchema);
export default Candidate;
