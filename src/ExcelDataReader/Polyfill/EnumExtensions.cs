#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
namespace System;

internal static class EnumExtensions
{
    extension(Enum)
    {
        public static TEnum Parse<TEnum>(string value)
            where TEnum : struct, Enum =>
            (TEnum)Enum.Parse(typeof(TEnum), value);
    }
}
#endif
