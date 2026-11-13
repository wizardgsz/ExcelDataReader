#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
namespace System.Collections.Generic;

internal static class DictionaryExtensions
{
    public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue value)
        where TKey : notnull
    {
        if (dictionary.ContainsKey(key))
            return false;

        dictionary.Add(key, value);
        return true;
    }
}
#endif
