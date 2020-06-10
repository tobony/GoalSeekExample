﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace DeliveryDate
{
    public class GoalSeek
    {
        public const decimal DefaultAccuracyLevel = 0.0000001m;
        public const int DefaultMaxIterations = 25;
        public const bool DefaultResultRoundOff = true;

        [System.ComponentModel.DefaultValue(DefaultMaxIterations)]
        public int MaxIterations { get; set; }
        [System.ComponentModel.DefaultValue(DefaultResultRoundOff)]
        public bool ResultRoundOff { get; set; }
        [System.ComponentModel.DefaultValue(DefaultAccuracyLevel)]
        public List<decimal> AccuracyLevels { get; }

        private readonly Func<decimal, decimal> func;

        public GoalSeek(IGoalSeek iGoalSeek)
        {
            func = iGoalSeek.Calculate;
        }

        public GoalSeekResult TrySeek(decimal targetValue = 0, decimal guess = 0)
        {
            return TrySeek(func, AccuracyLevels, targetValue, guess, MaxIterations, ResultRoundOff);
        }

        public static GoalSeekResult TrySeek(Func<decimal, decimal> func, List<decimal> accuracyLevels, decimal targetValue = 0, decimal guess = 0, int maxIterations = DefaultMaxIterations, bool resultRoundOff = DefaultResultRoundOff)
        {
            if (accuracyLevels.Count == 0) {
                throw new Exception("There should be at lease one accuracy level");
            }

            accuracyLevels = accuracyLevels.OrderBy(o => o).ToList();

            var iterations = 0;
            for (var i = 0; i < accuracyLevels.Count; i++)
            {
                var accuracyLevel = accuracyLevels[i];

                var goalSeekResult = TrySeek(func, accuracyLevel, targetValue, guess, maxIterations, resultRoundOff);

                iterations += goalSeekResult.Iterations;

                if (goalSeekResult.IsGoalReached || i == (accuracyLevels.Count - 1))
                    return new GoalSeekResult(goalSeekResult.TargetValue, goalSeekResult.AccucracyLevel, iterations, goalSeekResult.IsGoalReached, goalSeekResult.ClosestValue);
            }

            return null;
        }

        public static GoalSeekResult TrySeek(Func<decimal, decimal> func, decimal accuracyLevel = DefaultAccuracyLevel, decimal targetValue = 0, decimal guess = 0, int maxIterations = DefaultMaxIterations, bool resultRoundOff = DefaultResultRoundOff)
        {
            const decimal delta = 0.0001m;

            var iterations = 0;

            var result1 = (func(guess) - targetValue);
            while (Math.Abs(result1) > accuracyLevel && iterations++ < maxIterations)
            {
                var newGuess = (guess + delta);
                var result2 = (func(newGuess) - targetValue);
                if ((result2 - result1) != 0)
                    guess = guess - result1 * (newGuess - guess) / (result2 - result1);
                else
                    break;

                result1 = (func(guess) - targetValue);
            }

            if (iterations > maxIterations)
                iterations = maxIterations;

            if (resultRoundOff)
                guess = Math.Round(guess, (accuracyLevel.ToString().Length - (accuracyLevel.ToString().IndexOf('.') + 1)));

            return new GoalSeekResult(targetValue: targetValue, accucracyLevel: accuracyLevel, iterations: iterations, isGoalReached: Math.Abs(result1) <= accuracyLevel, closestValue: guess);
        }
    }
}
