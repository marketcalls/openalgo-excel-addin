/* global clearInterval, console, setInterval */
import { OpenAlgoAPI } from '../openalgoApi';
import { Settings } from '../settings';

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
export function logMessage(message) {
  console.log(message);

  return message;
}

/**
 * Fetches funds and margin details from the trading account
 * @customfunction FUNDS
 * @returns {string[][]} Returns a 2D array containing fund details
 */
export async function funds() {
    try {
        const settings = Settings.getSettings();
        const api = new OpenAlgoAPI(settings.host, settings.apikey);
        const response = await api.getFunds();
        
        if (response.status === 'success') {
            const fundData = response.data;
            return [
                ['Available Cash', fundData.availablecash],
                ['Collateral', fundData.collateral],
                ['M2M Realized', fundData.m2mrealized],
                ['M2M Unrealized', fundData.m2munrealized],
                ['Utilised Debits', fundData.utiliseddebits]
            ];
        } else {
            throw new Error('Failed to fetch fund details');
        }
    } catch (error) {
        return [['Error', error.message]];
    }
}
